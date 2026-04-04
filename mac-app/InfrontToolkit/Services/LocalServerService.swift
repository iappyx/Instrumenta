// LocalServerService.swift
// Minimaler HTTP-Server auf Basis von Network.framework (NWListener).
// Serviert statische Dateien aus dem gebundeten addin-dist/ Verzeichnis.
//
// Sicherheitshinweise:
// - Lauscht ausschließlich auf 127.0.0.1 (loopback) – kein externer Zugriff
// - Path-Traversal-Prävention: ".." in Pfadkomponenten wird gefiltert
// - Kein TLS/HTTPS nötig: Office für Mac akzeptiert HTTP für localhost

import Foundation
import Network

actor LocalServerService {

    // MARK: Private state

    private var listener: NWListener?
    private var isRunning = false
    private var config: ServerConfig?

    // MARK: Public API

    func start(port: UInt16) throws {
        guard !isRunning else { return }

        let serverConfig = ServerConfig(port: Int(port))
        self.config = serverConfig

        guard serverConfig.isAddinDistPresent else {
            throw LocalServerError.addinDistMissing(serverConfig.addinDistURL.path)
        }

        let params = NWParameters.tcp
        params.allowLocalEndpointReuse = true

        guard let nwPort = NWEndpoint.Port(rawValue: port) else {
            throw LocalServerError.invalidPort(Int(port))
        }

        let newListener = try NWListener(using: params, on: nwPort)

        newListener.stateUpdateHandler = { state in
            if case .failed(let error) = state {
                print("[LocalServer] Listener failed: \(error)")
            }
        }

        newListener.newConnectionHandler = { [weak self] connection in
            Task {
                await self?.handleConnection(connection)
            }
        }

        newListener.start(queue: .global(qos: .userInitiated))
        self.listener = newListener
        self.isRunning = true
    }

    func stop() {
        listener?.cancel()
        listener = nil
        isRunning = false
        config = nil
    }

    // MARK: Connection handling

    private func handleConnection(_ connection: NWConnection) async {
        connection.start(queue: .global(qos: .userInitiated))

        let data = await receiveData(from: connection)
        guard let requestStr = data.flatMap({ String(data: $0, encoding: .utf8) }) else {
            connection.cancel()
            return
        }

        await respond(to: requestStr, connection: connection)
    }

    private func receiveData(from connection: NWConnection) async -> Data? {
        await withCheckedContinuation { continuation in
            connection.receive(
                minimumIncompleteLength: 1,
                maximumLength: 16_384
            ) { data, _, _, _ in
                continuation.resume(returning: data)
            }
        }
    }

    private func respond(to request: String, connection: NWConnection) async {
        guard let baseURL = config?.addinDistURL else {
            connection.cancel()
            return
        }

        let path = parsePath(from: request)
        let fileURL = resolveFileURL(path: path, base: baseURL)
        let responseData = buildResponse(for: fileURL)

        await withCheckedContinuation { continuation in
            connection.send(content: responseData, completion: .contentProcessed { _ in
                connection.cancel()
                continuation.resume()
            })
        }
    }

    // MARK: Request parsing

    private func parsePath(from request: String) -> String {
        // "GET /path?query HTTP/1.1"
        guard let firstLine = request.components(separatedBy: "\r\n").first else { return "/" }
        let parts = firstLine.components(separatedBy: " ")
        guard parts.count >= 2 else { return "/" }
        return parts[1].components(separatedBy: "?").first ?? "/"
    }

    private func resolveFileURL(path: String, base: URL) -> URL {
        var clean = path.hasPrefix("/") ? String(path.dropFirst()) : path
        if clean.isEmpty { clean = "taskpane.html" }

        // Path-Traversal-Prävention
        let safe = clean
            .components(separatedBy: "/")
            .filter { !$0.isEmpty && $0 != ".." && $0 != "." }
            .joined(separator: "/")

        return base.appendingPathComponent(safe)
    }

    // MARK: Response building

    private func buildResponse(for fileURL: URL) -> Data {
        guard FileManager.default.fileExists(atPath: fileURL.path),
              let body = try? Data(contentsOf: fileURL) else {
            return buildErrorResponse(status: 404, message: "Not Found")
        }

        let mime = mimeType(for: fileURL)
        let header = "HTTP/1.1 200 OK\r\n" +
                     "Content-Type: \(mime)\r\n" +
                     "Content-Length: \(body.count)\r\n" +
                     "Access-Control-Allow-Origin: *\r\n" +
                     "Cache-Control: no-cache\r\n" +
                     "\r\n"
        return Data(header.utf8) + body
    }

    private func buildErrorResponse(status: Int, message: String) -> Data {
        let body = "\(status) \(message)"
        let header = "HTTP/1.1 \(status) \(message)\r\n" +
                     "Content-Type: text/plain\r\n" +
                     "Content-Length: \(body.utf8.count)\r\n" +
                     "\r\n"
        return Data((header + body).utf8)
    }

    // MARK: MIME types

    private func mimeType(for url: URL) -> String {
        switch url.pathExtension.lowercased() {
        case "html":  return "text/html; charset=utf-8"
        case "js":    return "application/javascript; charset=utf-8"
        case "css":   return "text/css; charset=utf-8"
        case "json":  return "application/json; charset=utf-8"
        case "png":   return "image/png"
        case "jpg", "jpeg": return "image/jpeg"
        case "svg":   return "image/svg+xml"
        case "ico":   return "image/x-icon"
        case "woff2": return "font/woff2"
        case "woff":  return "font/woff"
        case "ttf":   return "font/ttf"
        default:      return "application/octet-stream"
        }
    }
}

// MARK: - Errors

enum LocalServerError: LocalizedError {
    case invalidPort(Int)
    case addinDistMissing(String)

    var errorDescription: String? {
        switch self {
        case .invalidPort(let p):
            return "Ungültiger Port: \(p). Erlaubt: 1025–65534."
        case .addinDistMissing(let path):
            return "Add-in-Assets fehlen unter: \(path)\n" +
                   "Bitte zuerst 'sh scripts/build-addin.sh' ausführen."
        }
    }
}
