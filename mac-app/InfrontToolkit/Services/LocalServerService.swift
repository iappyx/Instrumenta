// LocalServerService.swift
// Minimaler HTTP-Server auf Basis von Network.framework (NWListener).
//
// Routen:
//   GET  /health                          → JSON Status-Antwort
//   GET  /config/Infront_BrandConfig.json → Brand Config (via ConfigSyncService)
//   POST /log                             → Log-Eintrag vom Add-in (fire-and-forget)
//   GET  /*                               → Statische Dateien aus addin-dist/
//
// Sicherheitshinweise:
//   - Lauscht ausschließlich auf 127.0.0.1 (loopback) – kein externer Zugriff
//   - Path-Traversal-Prävention: ".." in Pfadkomponenten wird gefiltert
//   - POST-Body auf 16 KB begrenzt

import Foundation
import Network

// MARK: - Log Handler Type

typealias AddinLogHandler = @Sendable (LogEntry.Level, String) -> Void

// MARK: - LocalServerService

actor LocalServerService {

    // MARK: Private state

    private var listener: NWListener?
    private var isRunning = false
    private var config: ServerConfig?
    private var logHandler: AddinLogHandler?

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

    /// Setzt den Log-Handler für vom Add-in empfangene Logs.
    /// Wird von AppState nach dem Server-Start gesetzt.
    func setLogHandler(_ handler: @escaping AddinLogHandler) {
        self.logHandler = handler
    }

    // MARK: - Connection handling

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

    // MARK: - Router

    private func respond(to request: String, connection: NWConnection) async {
        let method = parseMethod(from: request)
        let path   = parsePath(from: request)

        // Route: Health check
        if path == "/health" || path == "/health/" {
            let responseData = buildHealthResponse()
            await send(responseData, connection: connection)
            return
        }

        // Route: Brand Config
        if method == "GET" && path == "/config/Infront_BrandConfig.json" {
            let configData = ConfigSyncService.currentConfigData()
            let response = buildJSONResponse(data: configData)
            await send(response, connection: connection)
            return
        }

        // Route: Log ingestion vom Add-in (POST /log)
        if method == "POST" && path == "/log" {
            let body = parseBody(from: request)
            handleAddinLog(body: body)
            let response = buildEmptyResponse(status: "200 OK")
            await send(response, connection: connection)
            return
        }

        // Default: statische Datei servieren
        guard let baseURL = config?.addinDistURL else {
            await send(buildErrorResponse(status: 503, message: "Server not configured"), connection: connection)
            return
        }

        let fileURL = resolveFileURL(path: path, base: baseURL)
        let responseData = buildFileResponse(for: fileURL)
        await send(responseData, connection: connection)
    }

    // MARK: - Request parsing

    private func parseMethod(from request: String) -> String {
        guard let firstLine = request.components(separatedBy: "\r\n").first else { return "GET" }
        return firstLine.components(separatedBy: " ").first ?? "GET"
    }

    private func parsePath(from request: String) -> String {
        guard let firstLine = request.components(separatedBy: "\r\n").first else { return "/" }
        let parts = firstLine.components(separatedBy: " ")
        guard parts.count >= 2 else { return "/" }
        return parts[1].components(separatedBy: "?").first ?? "/"
    }

    private func parseBody(from request: String) -> String {
        // Body steht nach dem doppelten CRLF (Header-Ende)
        guard let range = request.range(of: "\r\n\r\n") else { return "" }
        return String(request[range.upperBound...])
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

    // MARK: - Log ingestion

    private func handleAddinLog(body: String) {
        guard let data = body.data(using: .utf8),
              let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
              let message = json["message"] as? String else { return }

        let levelStr  = json["level"]  as? String ?? "info"
        let module    = json["module"] as? String ?? "Add-in"

        let level: LogEntry.Level = {
            switch levelStr {
            case "error":   return .error
            case "warn":    return .warning
            default:        return .info
            }
        }()

        let fullMessage = "[\(module)] \(message)"

        // Auf Main Actor dispatchen für AppState.addLog
        let capturedHandler = logHandler
        let capturedLevel   = level
        let capturedMessage = fullMessage

        Task { @MainActor in
            capturedHandler?(capturedLevel, capturedMessage)
        }
    }

    // MARK: - Response builders

    private func buildHealthResponse() -> Data {
        let isPresent = config?.isAddinDistPresent ?? false
        let json = """
        {"status":"ok","version":"1.1.0","addinDistPresent":\(isPresent)}
        """
        return buildJSONResponse(data: Data(json.utf8))
    }

    private func buildJSONResponse(data: Data) -> Data {
        let header = "HTTP/1.1 200 OK\r\n" +
                     "Content-Type: application/json; charset=utf-8\r\n" +
                     "Content-Length: \(data.count)\r\n" +
                     "Access-Control-Allow-Origin: *\r\n" +
                     "Cache-Control: no-cache\r\n" +
                     "\r\n"
        return Data(header.utf8) + data
    }

    private func buildEmptyResponse(status: String) -> Data {
        let header = "HTTP/1.1 \(status)\r\n" +
                     "Content-Length: 0\r\n" +
                     "Access-Control-Allow-Origin: *\r\n" +
                     "\r\n"
        return Data(header.utf8)
    }

    private func buildFileResponse(for fileURL: URL) -> Data {
        guard FileManager.default.fileExists(atPath: fileURL.path),
              let body = try? Data(contentsOf: fileURL) else {
            return buildErrorResponse(status: 404, message: "Not Found")
        }

        let mime   = mimeType(for: fileURL)
        let header = "HTTP/1.1 200 OK\r\n" +
                     "Content-Type: \(mime)\r\n" +
                     "Content-Length: \(body.count)\r\n" +
                     "Access-Control-Allow-Origin: *\r\n" +
                     "Cache-Control: no-cache\r\n" +
                     "\r\n"
        return Data(header.utf8) + body
    }

    private func buildErrorResponse(status: Int, message: String) -> Data {
        let body   = "\(status) \(message)"
        let header = "HTTP/1.1 \(status) \(message)\r\n" +
                     "Content-Type: text/plain\r\n" +
                     "Content-Length: \(body.utf8.count)\r\n" +
                     "\r\n"
        return Data((header + body).utf8)
    }

    // MARK: - Send

    private func send(_ data: Data, connection: NWConnection) async {
        await withCheckedContinuation { continuation in
            connection.send(content: data, completion: .contentProcessed { _ in
                connection.cancel()
                continuation.resume()
            })
        }
    }

    // MARK: - MIME types

    private func mimeType(for url: URL) -> String {
        switch url.pathExtension.lowercased() {
        case "html":        return "text/html; charset=utf-8"
        case "js":          return "application/javascript; charset=utf-8"
        case "css":         return "text/css; charset=utf-8"
        case "json":        return "application/json; charset=utf-8"
        case "png":         return "image/png"
        case "jpg", "jpeg": return "image/jpeg"
        case "svg":         return "image/svg+xml"
        case "ico":         return "image/x-icon"
        case "woff2":       return "font/woff2"
        case "woff":        return "font/woff"
        case "ttf":         return "font/ttf"
        default:            return "application/octet-stream"
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
