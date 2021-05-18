Attribute VB_Name = "ModuleSpellcheckLanguage"
'MIT License

'Copyright (c) 2021 iappyx

'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.


Sub ShowChangeSpellCheckLanguageForm()
    
    Dim LanguageNames(1 To 216) As String
    LanguageNames(1) = "Afrikaans"
    LanguageNames(2) = "Albanian"
    LanguageNames(3) = "Amharic"
    LanguageNames(4) = "Arabic"
    LanguageNames(5) = "Arabic Algeria"
    LanguageNames(6) = "Arabic Bahrain"
    LanguageNames(7) = "Arabic Egypt"
    LanguageNames(8) = "Arabic Iraq"
    LanguageNames(9) = "Arabic Jordan"
    LanguageNames(10) = "Arabic Kuwait"
    LanguageNames(11) = "Arabic Lebanon"
    LanguageNames(12) = "Arabic Libya"
    LanguageNames(13) = "Arabic Morocco"
    LanguageNames(14) = "Arabic Oman"
    LanguageNames(15) = "Arabic Qatar"
    LanguageNames(16) = "Arabic Syria"
    LanguageNames(17) = "Arabic Tunisia"
    LanguageNames(18) = "Arabic UAE"
    LanguageNames(19) = "Arabic Yemen"
    LanguageNames(20) = "Armenian"
    LanguageNames(21) = "Assamese"
    LanguageNames(22) = "Azerbaijani Cyrillic"
    LanguageNames(23) = "Azerbaijani Latin"
    LanguageNames(24) = "Basque (Basque)"
    LanguageNames(25) = "Belgian Dutch"
    LanguageNames(26) = "Belgian French"
    LanguageNames(27) = "Bengali"
    LanguageNames(28) = "Bosnian"
    LanguageNames(29) = "Bosnian Bosnia Herzegovina Cyrillic"
    LanguageNames(30) = "Bosnian Bosnia Herzegovina Latin"
    LanguageNames(31) = "Portuguese (Brazil)"
    LanguageNames(32) = "Bulgarian"
    LanguageNames(33) = "Burmese"
    LanguageNames(34) = "Belarusian"
    LanguageNames(35) = "Catalan"
    LanguageNames(36) = "Cherokee"
    LanguageNames(37) = "Chinese Hong Kong SAR"
    LanguageNames(38) = "Chinese Macao SAR"
    LanguageNames(39) = "Chinese Singapore"
    LanguageNames(40) = "Croatian"
    LanguageNames(41) = "Czech"
    LanguageNames(42) = "Danish"
    LanguageNames(43) = "Divehi"
    LanguageNames(44) = "Dutch"
    LanguageNames(45) = "Edo"
    LanguageNames(46) = "English AUS"
    LanguageNames(47) = "English Belize"
    LanguageNames(48) = "English Canadian"
    LanguageNames(49) = "English Caribbean"
    LanguageNames(50) = "English Indonesia"
    LanguageNames(51) = "English Ireland"
    LanguageNames(52) = "English Jamaica"
    LanguageNames(53) = "English NewZealand"
    LanguageNames(54) = "English Philippines"
    LanguageNames(55) = "English South Africa"
    LanguageNames(56) = "English Trinidad Tobago"
    LanguageNames(57) = "English UK"
    LanguageNames(58) = "English US"
    LanguageNames(59) = "English Zimbabwe"
    LanguageNames(60) = "Estonian"
    LanguageNames(61) = "Faeroese"
    LanguageNames(62) = "Farsi"
    LanguageNames(63) = "Filipino"
    LanguageNames(64) = "Finnish"
    LanguageNames(65) = "French"
    LanguageNames(66) = "French Cameroon"
    LanguageNames(67) = "French Canadian"
    LanguageNames(68) = "French Coted Ivoire"
    LanguageNames(69) = "French Haiti"
    LanguageNames(70) = "French Luxembourg"
    LanguageNames(71) = "French Mali"
    LanguageNames(72) = "French Monaco"
    LanguageNames(73) = "French Morocco"
    LanguageNames(74) = "French Reunion"
    LanguageNames(75) = "French Senegal"
    LanguageNames(76) = "French West Indies"
    LanguageNames(77) = "French Congo DRC"
    LanguageNames(78) = "Frisian Netherlands"
    LanguageNames(79) = "Fulfulde"
    LanguageNames(80) = "Irish (Ireland)"
    LanguageNames(81) = "Scottish Gaelic"
    LanguageNames(82) = "Galician"
    LanguageNames(83) = "Georgian"
    LanguageNames(84) = "German"
    LanguageNames(85) = "German Austria"
    LanguageNames(86) = "German Liechtenstein"
    LanguageNames(87) = "German Luxembourg"
    LanguageNames(88) = "Greek"
    LanguageNames(89) = "Guarani"
    LanguageNames(90) = "Gujarati"
    LanguageNames(91) = "Hausa"
    LanguageNames(92) = "Hawaiian"
    LanguageNames(93) = "Hebrew"
    LanguageNames(94) = "Hindi"
    LanguageNames(95) = "Hungarian"
    LanguageNames(96) = "Ibibio"
    LanguageNames(97) = "Icelandic"
    LanguageNames(98) = "Igbo"
    LanguageNames(99) = "Indonesian"
    LanguageNames(100) = "Inuktitut"
    LanguageNames(101) = "Italian"
    LanguageNames(102) = "Japanese"
    LanguageNames(103) = "Kannada"
    LanguageNames(104) = "Kanuri"
    LanguageNames(105) = "Kashmiri"
    LanguageNames(106) = "Kashmiri Devanagari"
    LanguageNames(107) = "Kazakh"
    LanguageNames(108) = "Khmer"
    LanguageNames(109) = "Kirghiz"
    LanguageNames(110) = "Konkani"
    LanguageNames(111) = "Korean"
    LanguageNames(112) = "Kyrgyz"
    LanguageNames(113) = "Lao"
    LanguageNames(114) = "Latin"
    LanguageNames(115) = "Latvian"
    LanguageNames(116) = "Lithuanian"
    LanguageNames(117) = "Macedonian FYROM"
    LanguageNames(118) = "Malayalam"
    LanguageNames(119) = "Malay Brunei Darussalam"
    LanguageNames(120) = "Malaysian"
    LanguageNames(121) = "Maltese"
    LanguageNames(122) = "Manipuri"
    LanguageNames(123) = "Maori"
    LanguageNames(124) = "Marathi"
    LanguageNames(125) = "Mexican Spanish"
    LanguageNames(126) = "Mixed"
    LanguageNames(127) = "Mongolian"
    LanguageNames(128) = "Nepali"
    LanguageNames(129) = "No specified"
    LanguageNames(130) = "No proofing"
    LanguageNames(131) = "Norwegian Bokmol"
    LanguageNames(132) = "Norwegian Nynorsk"
    LanguageNames(133) = "Odia"
    LanguageNames(134) = "Oromo"
    LanguageNames(135) = "Pashto"
    LanguageNames(136) = "Polish"
    LanguageNames(137) = "Portuguese"
    LanguageNames(138) = "Punjabi"
    LanguageNames(139) = "Quechua Bolivia"
    LanguageNames(140) = "Quechua Ecuador"
    LanguageNames(141) = "Quechua Peru"
    LanguageNames(142) = "Rhaeto Romanic"
    LanguageNames(143) = "Romanian"
    LanguageNames(144) = "Romanian Moldova"
    LanguageNames(145) = "Russian"
    LanguageNames(146) = "Russian Moldova"
    LanguageNames(147) = "Sami Lappish"
    LanguageNames(148) = "Sanskrit"
    LanguageNames(149) = "Sepedi"
    LanguageNames(150) = "Serbian Bosnia Herzegovina Cyrillic"
    LanguageNames(151) = "Serbian Bosnia Herzegovina Latin"
    LanguageNames(152) = "Serbian Cyrillic"
    LanguageNames(153) = "Serbian Latin"
    LanguageNames(154) = "Sesotho"
    LanguageNames(155) = "Simplified Chinese"
    LanguageNames(156) = "Sindhi"
    LanguageNames(157) = "Sindhi Pakistan"
    LanguageNames(158) = "Sinhalese"
    LanguageNames(159) = "Slovak"
    LanguageNames(160) = "Slovenian"
    LanguageNames(161) = "Somali"
    LanguageNames(162) = "Sorbian"
    LanguageNames(163) = "Spanish"
    LanguageNames(164) = "Spanish Argentina"
    LanguageNames(165) = "Spanish Bolivia"
    LanguageNames(166) = "Spanish Chile"
    LanguageNames(167) = "Spanish Colombia"
    LanguageNames(168) = "Spanish Costa Rica"
    LanguageNames(169) = "Spanish Dominican Republic"
    LanguageNames(170) = "Spanish Ecuador"
    LanguageNames(171) = "Spanish El Salvador"
    LanguageNames(172) = "Spanish Guatemala"
    LanguageNames(173) = "Spanish Honduras"
    LanguageNames(174) = "Spanish Modern Sort"
    LanguageNames(175) = "Spanish Nicaragua"
    LanguageNames(176) = "Spanish Panama"
    LanguageNames(177) = "Spanish Paraguay"
    LanguageNames(178) = "Spanish Peru"
    LanguageNames(179) = "Spanish Puerto Rico"
    LanguageNames(180) = "Spanish Uruguay"
    LanguageNames(181) = "Spanish Venezuela"
    LanguageNames(182) = "Sutu"
    LanguageNames(183) = "Swahili"
    LanguageNames(184) = "Swedish"
    LanguageNames(185) = "Swedish Finland"
    LanguageNames(186) = "Swiss French"
    LanguageNames(187) = "Swiss German"
    LanguageNames(188) = "Swiss Italian"
    LanguageNames(189) = "Syriac"
    LanguageNames(190) = "Tajik"
    LanguageNames(191) = "Tamazight"
    LanguageNames(192) = "Tamazight Latin"
    LanguageNames(193) = "Tamil"
    LanguageNames(194) = "Tatar"
    LanguageNames(195) = "Telugu"
    LanguageNames(196) = "Thai"
    LanguageNames(197) = "Tibetan"
    LanguageNames(198) = "Tigrigna Eritrea"
    LanguageNames(199) = "Tigrigna Ethiopic"
    LanguageNames(200) = "Traditional Chinese"
    LanguageNames(201) = "Tsonga"
    LanguageNames(202) = "Tswana"
    LanguageNames(203) = "Turkish"
    LanguageNames(204) = "Turkmen"
    LanguageNames(205) = "Ukrainian"
    LanguageNames(206) = "Urdu"
    LanguageNames(207) = "Uzbek Cyrillic"
    LanguageNames(208) = "Uzbek Latin"
    LanguageNames(209) = "Venda"
    LanguageNames(210) = "Vietnamese"
    LanguageNames(211) = "Welsh"
    LanguageNames(212) = "Xhosa"
    LanguageNames(213) = "Yi"
    LanguageNames(214) = "Yiddish"
    LanguageNames(215) = "Yoruba"
    LanguageNames(216) = "Zulu"
    
    ChangeSpellCheckLanguageForm.ComboBox1.Clear
    For I = 1 To 216
        ChangeSpellCheckLanguageForm.ComboBox1.AddItem LanguageNames(I)
    Next
    
    ChangeSpellCheckLanguageForm.Show
    
End Sub

Sub ChangeSpellCheckLanguage()
    
    Dim LanguageNames(1 To 216) As String
    LanguageNames(1) = "Afrikaans"
    LanguageNames(2) = "Albanian"
    LanguageNames(3) = "Amharic"
    LanguageNames(4) = "Arabic"
    LanguageNames(5) = "Arabic Algeria"
    LanguageNames(6) = "Arabic Bahrain"
    LanguageNames(7) = "Arabic Egypt"
    LanguageNames(8) = "Arabic Iraq"
    LanguageNames(9) = "Arabic Jordan"
    LanguageNames(10) = "Arabic Kuwait"
    LanguageNames(11) = "Arabic Lebanon"
    LanguageNames(12) = "Arabic Libya"
    LanguageNames(13) = "Arabic Morocco"
    LanguageNames(14) = "Arabic Oman"
    LanguageNames(15) = "Arabic Qatar"
    LanguageNames(16) = "Arabic Syria"
    LanguageNames(17) = "Arabic Tunisia"
    LanguageNames(18) = "Arabic UAE"
    LanguageNames(19) = "Arabic Yemen"
    LanguageNames(20) = "Armenian"
    LanguageNames(21) = "Assamese"
    LanguageNames(22) = "Azerbaijani Cyrillic"
    LanguageNames(23) = "Azerbaijani Latin"
    LanguageNames(24) = "Basque (Basque)"
    LanguageNames(25) = "Belgian Dutch"
    LanguageNames(26) = "Belgian French"
    LanguageNames(27) = "Bengali"
    LanguageNames(28) = "Bosnian"
    LanguageNames(29) = "Bosnian Bosnia Herzegovina Cyrillic"
    LanguageNames(30) = "Bosnian Bosnia Herzegovina Latin"
    LanguageNames(31) = "Portuguese (Brazil)"
    LanguageNames(32) = "Bulgarian"
    LanguageNames(33) = "Burmese"
    LanguageNames(34) = "Belarusian"
    LanguageNames(35) = "Catalan"
    LanguageNames(36) = "Cherokee"
    LanguageNames(37) = "Chinese Hong Kong SAR"
    LanguageNames(38) = "Chinese Macao SAR"
    LanguageNames(39) = "Chinese Singapore"
    LanguageNames(40) = "Croatian"
    LanguageNames(41) = "Czech"
    LanguageNames(42) = "Danish"
    LanguageNames(43) = "Divehi"
    LanguageNames(44) = "Dutch"
    LanguageNames(45) = "Edo"
    LanguageNames(46) = "English AUS"
    LanguageNames(47) = "English Belize"
    LanguageNames(48) = "English Canadian"
    LanguageNames(49) = "English Caribbean"
    LanguageNames(50) = "English Indonesia"
    LanguageNames(51) = "English Ireland"
    LanguageNames(52) = "English Jamaica"
    LanguageNames(53) = "English NewZealand"
    LanguageNames(54) = "English Philippines"
    LanguageNames(55) = "English South Africa"
    LanguageNames(56) = "English Trinidad Tobago"
    LanguageNames(57) = "English UK"
    LanguageNames(58) = "English US"
    LanguageNames(59) = "English Zimbabwe"
    LanguageNames(60) = "Estonian"
    LanguageNames(61) = "Faeroese"
    LanguageNames(62) = "Farsi"
    LanguageNames(63) = "Filipino"
    LanguageNames(64) = "Finnish"
    LanguageNames(65) = "French"
    LanguageNames(66) = "French Cameroon"
    LanguageNames(67) = "French Canadian"
    LanguageNames(68) = "French Coted Ivoire"
    LanguageNames(69) = "French Haiti"
    LanguageNames(70) = "French Luxembourg"
    LanguageNames(71) = "French Mali"
    LanguageNames(72) = "French Monaco"
    LanguageNames(73) = "French Morocco"
    LanguageNames(74) = "French Reunion"
    LanguageNames(75) = "French Senegal"
    LanguageNames(76) = "French West Indies"
    LanguageNames(77) = "French Congo DRC"
    LanguageNames(78) = "Frisian Netherlands"
    LanguageNames(79) = "Fulfulde"
    LanguageNames(80) = "Irish (Ireland)"
    LanguageNames(81) = "Scottish Gaelic"
    LanguageNames(82) = "Galician"
    LanguageNames(83) = "Georgian"
    LanguageNames(84) = "German"
    LanguageNames(85) = "German Austria"
    LanguageNames(86) = "German Liechtenstein"
    LanguageNames(87) = "German Luxembourg"
    LanguageNames(88) = "Greek"
    LanguageNames(89) = "Guarani"
    LanguageNames(90) = "Gujarati"
    LanguageNames(91) = "Hausa"
    LanguageNames(92) = "Hawaiian"
    LanguageNames(93) = "Hebrew"
    LanguageNames(94) = "Hindi"
    LanguageNames(95) = "Hungarian"
    LanguageNames(96) = "Ibibio"
    LanguageNames(97) = "Icelandic"
    LanguageNames(98) = "Igbo"
    LanguageNames(99) = "Indonesian"
    LanguageNames(100) = "Inuktitut"
    LanguageNames(101) = "Italian"
    LanguageNames(102) = "Japanese"
    LanguageNames(103) = "Kannada"
    LanguageNames(104) = "Kanuri"
    LanguageNames(105) = "Kashmiri"
    LanguageNames(106) = "Kashmiri Devanagari"
    LanguageNames(107) = "Kazakh"
    LanguageNames(108) = "Khmer"
    LanguageNames(109) = "Kirghiz"
    LanguageNames(110) = "Konkani"
    LanguageNames(111) = "Korean"
    LanguageNames(112) = "Kyrgyz"
    LanguageNames(113) = "Lao"
    LanguageNames(114) = "Latin"
    LanguageNames(115) = "Latvian"
    LanguageNames(116) = "Lithuanian"
    LanguageNames(117) = "Macedonian FYROM"
    LanguageNames(118) = "Malayalam"
    LanguageNames(119) = "Malay Brunei Darussalam"
    LanguageNames(120) = "Malaysian"
    LanguageNames(121) = "Maltese"
    LanguageNames(122) = "Manipuri"
    LanguageNames(123) = "Maori"
    LanguageNames(124) = "Marathi"
    LanguageNames(125) = "Mexican Spanish"
    LanguageNames(126) = "Mixed"
    LanguageNames(127) = "Mongolian"
    LanguageNames(128) = "Nepali"
    LanguageNames(129) = "No specified"
    LanguageNames(130) = "No proofing"
    LanguageNames(131) = "Norwegian Bokmol"
    LanguageNames(132) = "Norwegian Nynorsk"
    LanguageNames(133) = "Odia"
    LanguageNames(134) = "Oromo"
    LanguageNames(135) = "Pashto"
    LanguageNames(136) = "Polish"
    LanguageNames(137) = "Portuguese"
    LanguageNames(138) = "Punjabi"
    LanguageNames(139) = "Quechua Bolivia"
    LanguageNames(140) = "Quechua Ecuador"
    LanguageNames(141) = "Quechua Peru"
    LanguageNames(142) = "Rhaeto Romanic"
    LanguageNames(143) = "Romanian"
    LanguageNames(144) = "Romanian Moldova"
    LanguageNames(145) = "Russian"
    LanguageNames(146) = "Russian Moldova"
    LanguageNames(147) = "Sami Lappish"
    LanguageNames(148) = "Sanskrit"
    LanguageNames(149) = "Sepedi"
    LanguageNames(150) = "Serbian Bosnia Herzegovina Cyrillic"
    LanguageNames(151) = "Serbian Bosnia Herzegovina Latin"
    LanguageNames(152) = "Serbian Cyrillic"
    LanguageNames(153) = "Serbian Latin"
    LanguageNames(154) = "Sesotho"
    LanguageNames(155) = "Simplified Chinese"
    LanguageNames(156) = "Sindhi"
    LanguageNames(157) = "Sindhi Pakistan"
    LanguageNames(158) = "Sinhalese"
    LanguageNames(159) = "Slovak"
    LanguageNames(160) = "Slovenian"
    LanguageNames(161) = "Somali"
    LanguageNames(162) = "Sorbian"
    LanguageNames(163) = "Spanish"
    LanguageNames(164) = "Spanish Argentina"
    LanguageNames(165) = "Spanish Bolivia"
    LanguageNames(166) = "Spanish Chile"
    LanguageNames(167) = "Spanish Colombia"
    LanguageNames(168) = "Spanish Costa Rica"
    LanguageNames(169) = "Spanish Dominican Republic"
    LanguageNames(170) = "Spanish Ecuador"
    LanguageNames(171) = "Spanish El Salvador"
    LanguageNames(172) = "Spanish Guatemala"
    LanguageNames(173) = "Spanish Honduras"
    LanguageNames(174) = "Spanish Modern Sort"
    LanguageNames(175) = "Spanish Nicaragua"
    LanguageNames(176) = "Spanish Panama"
    LanguageNames(177) = "Spanish Paraguay"
    LanguageNames(178) = "Spanish Peru"
    LanguageNames(179) = "Spanish Puerto Rico"
    LanguageNames(180) = "Spanish Uruguay"
    LanguageNames(181) = "Spanish Venezuela"
    LanguageNames(182) = "Sutu"
    LanguageNames(183) = "Swahili"
    LanguageNames(184) = "Swedish"
    LanguageNames(185) = "Swedish Finland"
    LanguageNames(186) = "Swiss French"
    LanguageNames(187) = "Swiss German"
    LanguageNames(188) = "Swiss Italian"
    LanguageNames(189) = "Syriac"
    LanguageNames(190) = "Tajik"
    LanguageNames(191) = "Tamazight"
    LanguageNames(192) = "Tamazight Latin"
    LanguageNames(193) = "Tamil"
    LanguageNames(194) = "Tatar"
    LanguageNames(195) = "Telugu"
    LanguageNames(196) = "Thai"
    LanguageNames(197) = "Tibetan"
    LanguageNames(198) = "Tigrigna Eritrea"
    LanguageNames(199) = "Tigrigna Ethiopic"
    LanguageNames(200) = "Traditional Chinese"
    LanguageNames(201) = "Tsonga"
    LanguageNames(202) = "Tswana"
    LanguageNames(203) = "Turkish"
    LanguageNames(204) = "Turkmen"
    LanguageNames(205) = "Ukrainian"
    LanguageNames(206) = "Urdu"
    LanguageNames(207) = "Uzbek Cyrillic"
    LanguageNames(208) = "Uzbek Latin"
    LanguageNames(209) = "Venda"
    LanguageNames(210) = "Vietnamese"
    LanguageNames(211) = "Welsh"
    LanguageNames(212) = "Xhosa"
    LanguageNames(213) = "Yi"
    LanguageNames(214) = "Yiddish"
    LanguageNames(215) = "Yoruba"
    LanguageNames(216) = "Zulu"
    
    Dim LanguageIDs(1 To 216) As String
    LanguageIDs(1) = msoLanguageIDAfrikaans
    LanguageIDs(2) = msoLanguageIDAlbanian
    LanguageIDs(3) = msoLanguageIDAmharic
    LanguageIDs(4) = msoLanguageIDArabic
    LanguageIDs(5) = msoLanguageIDArabicAlgeria
    LanguageIDs(6) = msoLanguageIDArabicBahrain
    LanguageIDs(7) = msoLanguageIDArabicEgypt
    LanguageIDs(8) = msoLanguageIDArabicIraq
    LanguageIDs(9) = msoLanguageIDArabicJordan
    LanguageIDs(10) = msoLanguageIDArabicKuwait
    LanguageIDs(11) = msoLanguageIDArabicLebanon
    LanguageIDs(12) = msoLanguageIDArabicLibya
    LanguageIDs(13) = msoLanguageIDArabicMorocco
    LanguageIDs(14) = msoLanguageIDArabicOman
    LanguageIDs(15) = msoLanguageIDArabicQatar
    LanguageIDs(16) = msoLanguageIDArabicSyria
    LanguageIDs(17) = msoLanguageIDArabicTunisia
    LanguageIDs(18) = msoLanguageIDArabicUAE
    LanguageIDs(19) = msoLanguageIDArabicYemen
    LanguageIDs(20) = msoLanguageIDArmenian
    LanguageIDs(21) = msoLanguageIDAssamese
    LanguageIDs(22) = msoLanguageIDAzeriCyrillic
    LanguageIDs(23) = msoLanguageIDAzeriLatin
    LanguageIDs(24) = msoLanguageIDBasque
    LanguageIDs(25) = msoLanguageIDBelgianDutch
    LanguageIDs(26) = msoLanguageIDBelgianFrench
    LanguageIDs(27) = msoLanguageIDBengali
    LanguageIDs(28) = msoLanguageIDBosnian
    LanguageIDs(29) = msoLanguageIDBosnianBosniaHerzegovinaCyrillic
    LanguageIDs(30) = msoLanguageIDBosnianBosniaHerzegovinaLatin
    LanguageIDs(31) = msoLanguageIDBrazilianPortuguese
    LanguageIDs(32) = msoLanguageIDBulgarian
    LanguageIDs(33) = msoLanguageIDBurmese
    LanguageIDs(34) = msoLanguageIDByelorussian
    LanguageIDs(35) = msoLanguageIDCatalan
    LanguageIDs(36) = msoLanguageIDCherokee
    LanguageIDs(37) = msoLanguageIDChineseHongKongSAR
    LanguageIDs(38) = msoLanguageIDChineseMacaoSAR
    LanguageIDs(39) = msoLanguageIDChineseSingapore
    LanguageIDs(40) = msoLanguageIDCroatian
    LanguageIDs(41) = msoLanguageIDCzech
    LanguageIDs(42) = msoLanguageIDDanish
    LanguageIDs(43) = msoLanguageIDDivehi
    LanguageIDs(44) = msoLanguageIDDutch
    LanguageIDs(45) = msoLanguageIDEdo
    LanguageIDs(46) = msoLanguageIDEnglishAUS
    LanguageIDs(47) = msoLanguageIDEnglishBelize
    LanguageIDs(48) = msoLanguageIDEnglishCanadian
    LanguageIDs(49) = msoLanguageIDEnglishCaribbean
    LanguageIDs(50) = msoLanguageIDEnglishIndonesia
    LanguageIDs(51) = msoLanguageIDEnglishIreland
    LanguageIDs(52) = msoLanguageIDEnglishJamaica
    LanguageIDs(53) = msoLanguageIDEnglishNewZealand
    LanguageIDs(54) = msoLanguageIDEnglishPhilippines
    LanguageIDs(55) = msoLanguageIDEnglishSouthAfrica
    LanguageIDs(56) = msoLanguageIDEnglishTrinidadTobago
    LanguageIDs(57) = msoLanguageIDEnglishUK
    LanguageIDs(58) = msoLanguageIDEnglishUS
    LanguageIDs(59) = msoLanguageIDEnglishZimbabwe
    LanguageIDs(60) = msoLanguageIDEstonian
    LanguageIDs(61) = msoLanguageIDFaeroese
    LanguageIDs(62) = msoLanguageIDFarsi
    LanguageIDs(63) = msoLanguageIDFilipino
    LanguageIDs(64) = msoLanguageIDFinnish
    LanguageIDs(65) = msoLanguageIDFrench
    LanguageIDs(66) = msoLanguageIDFrenchCameroon
    LanguageIDs(67) = msoLanguageIDFrenchCanadian
    LanguageIDs(68) = msoLanguageIDFrenchCotedIvoire
    LanguageIDs(69) = msoLanguageIDFrenchHaiti
    LanguageIDs(70) = msoLanguageIDFrenchLuxembourg
    LanguageIDs(71) = msoLanguageIDFrenchMali
    LanguageIDs(72) = msoLanguageIDFrenchMonaco
    LanguageIDs(73) = msoLanguageIDFrenchMorocco
    LanguageIDs(74) = msoLanguageIDFrenchReunion
    LanguageIDs(75) = msoLanguageIDFrenchSenegal
    LanguageIDs(76) = msoLanguageIDFrenchWestIndies
    LanguageIDs(77) = msoLanguageIDFranchCongoDRC
    LanguageIDs(78) = msoLanguageIDFrisianNetherlands
    LanguageIDs(79) = msoLanguageIDFulfulde
    LanguageIDs(80) = msoLanguageIDGaelicIreland
    LanguageIDs(81) = msoLanguageIDGaelicScotland
    LanguageIDs(82) = msoLanguageIDGalician
    LanguageIDs(83) = msoLanguageIDGeorgian
    LanguageIDs(84) = msoLanguageIDGerman
    LanguageIDs(85) = msoLanguageIDGermanAustria
    LanguageIDs(86) = msoLanguageIDGermanLiechtenstein
    LanguageIDs(87) = msoLanguageIDGermanLuxembourg
    LanguageIDs(88) = msoLanguageIDGreek
    LanguageIDs(89) = msoLanguageIDGuarani
    LanguageIDs(90) = msoLanguageIDGujarati
    LanguageIDs(91) = msoLanguageIDHausa
    LanguageIDs(92) = msoLanguageIDHawaiian
    LanguageIDs(93) = msoLanguageIDHebrew
    LanguageIDs(94) = msoLanguageIDHindi
    LanguageIDs(95) = msoLanguageIDHungarian
    LanguageIDs(96) = msoLanguageIDIbibio
    LanguageIDs(97) = msoLanguageIDIcelandic
    LanguageIDs(98) = msoLanguageIDIgbo
    LanguageIDs(99) = msoLanguageIDIndonesian
    LanguageIDs(100) = msoLanguageIDInuktitut
    LanguageIDs(101) = msoLanguageIDItalian
    LanguageIDs(102) = msoLanguageIDJapanese
    LanguageIDs(103) = msoLanguageIDKannada
    LanguageIDs(104) = msoLanguageIDKanuri
    LanguageIDs(105) = msoLanguageIDKashmiri
    LanguageIDs(106) = msoLanguageIDKashmiriDevanagari
    LanguageIDs(107) = msoLanguageIDKazakh
    LanguageIDs(108) = msoLanguageIDKhmer
    LanguageIDs(109) = msoLanguageIDKirghiz
    LanguageIDs(110) = msoLanguageIDKonkani
    LanguageIDs(111) = msoLanguageIDKorean
    LanguageIDs(112) = msoLanguageIDKyrgyz
    LanguageIDs(113) = msoLanguageIDLao
    LanguageIDs(114) = msoLanguageIDLatin
    LanguageIDs(115) = msoLanguageIDLatvian
    LanguageIDs(116) = msoLanguageIDLithuanian
    LanguageIDs(117) = msoLanguageIDMacedoninanFYROM
    LanguageIDs(118) = msoLanguageIDMalayalam
    LanguageIDs(119) = msoLanguageIDMalayBruneiDarussalam
    LanguageIDs(120) = msoLanguageIDMalaysian
    LanguageIDs(121) = msoLanguageIDMaltese
    LanguageIDs(122) = msoLanguageIDManipuri
    LanguageIDs(123) = msoLanguageIDMaori
    LanguageIDs(124) = msoLanguageIDMarathi
    LanguageIDs(125) = msoLanguageIDMexicanSpanish
    LanguageIDs(126) = msoLanguageIDMixed
    LanguageIDs(127) = msoLanguageIDMongolian
    LanguageIDs(128) = msoLanguageIDNepali
    LanguageIDs(129) = msoLanguageIDNone
    LanguageIDs(130) = msoLanguageIDNoProofing
    LanguageIDs(131) = msoLanguageIDNorwegianBokmol
    LanguageIDs(132) = msoLanguageIDNorwegianNynorsk
    LanguageIDs(133) = msoLanguageIDOriya
    LanguageIDs(134) = msoLanguageIDOromo
    LanguageIDs(135) = msoLanguageIDPashto
    LanguageIDs(136) = msoLanguageIDPolish
    LanguageIDs(137) = msoLanguageIDPortuguese
    LanguageIDs(138) = msoLanguageIDPunjabi
    LanguageIDs(139) = msoLanguageIDQuechuaBolivia
    LanguageIDs(140) = msoLanguageIDQuechuaEcuador
    LanguageIDs(141) = msoLanguageIDQuechuaPeru
    LanguageIDs(142) = msoLanguageIDRhaetoRomanic
    LanguageIDs(143) = msoLanguageIDRomanian
    LanguageIDs(144) = msoLanguageIDRomanianMoldova
    LanguageIDs(145) = msoLanguageIDRussian
    LanguageIDs(146) = msoLanguageIDRussianMoldova
    LanguageIDs(147) = msoLanguageIDSamiLappish
    LanguageIDs(148) = msoLanguageIDSanskrit
    LanguageIDs(149) = msoLanguageIDSepedi
    LanguageIDs(150) = msoLanguageIDSerbianBosniaHerzegovinaCyrillic
    LanguageIDs(151) = msoLanguageIDSerbianBosniaHerzegovinaLatin
    LanguageIDs(152) = msoLanguageIDSerbianCyrillic
    LanguageIDs(153) = msoLanguageIDSerbianLatin
    LanguageIDs(154) = msoLanguageIDSesotho
    LanguageIDs(155) = msoLanguageIDSimplifiedChinese
    LanguageIDs(156) = msoLanguageIDSindhi
    LanguageIDs(157) = msoLanguageIDSindhiPakistan
    LanguageIDs(158) = msoLanguageIDSinhalese
    LanguageIDs(159) = msoLanguageIDSlovak
    LanguageIDs(160) = msoLanguageIDSlovenian
    LanguageIDs(161) = msoLanguageIDSomali
    LanguageIDs(162) = msoLanguageIDSorbian
    LanguageIDs(163) = msoLanguageIDSpanish
    LanguageIDs(164) = msoLanguageIDSpanishArgentina
    LanguageIDs(165) = msoLanguageIDSpanishBolivia
    LanguageIDs(166) = msoLanguageIDSpanishChile
    LanguageIDs(167) = msoLanguageIDSpanishColombia
    LanguageIDs(168) = msoLanguageIDSpanishCostaRica
    LanguageIDs(169) = msoLanguageIDSpanishDominicanRepublic
    LanguageIDs(170) = msoLanguageIDSpanishEcuador
    LanguageIDs(171) = msoLanguageIDSpanishElSalvador
    LanguageIDs(172) = msoLanguageIDSpanishGuatemala
    LanguageIDs(173) = msoLanguageIDSpanishHonduras
    LanguageIDs(174) = msoLanguageIDSpanishModernSort
    LanguageIDs(175) = msoLanguageIDSpanishNicaragua
    LanguageIDs(176) = msoLanguageIDSpanishPanama
    LanguageIDs(177) = msoLanguageIDSpanishParaguay
    LanguageIDs(178) = msoLanguageIDSpanishPeru
    LanguageIDs(179) = msoLanguageIDSpanishPuertoRico
    LanguageIDs(180) = msoLanguageIDSpanishUruguay
    LanguageIDs(181) = msoLanguageIDSpanishVenezuela
    LanguageIDs(182) = msoLanguageIDSutu
    LanguageIDs(183) = msoLanguageIDSwahili
    LanguageIDs(184) = msoLanguageIDSwedish
    LanguageIDs(185) = msoLanguageIDSwedishFinland
    LanguageIDs(186) = msoLanguageIDSwissFrench
    LanguageIDs(187) = msoLanguageIDSwissGerman
    LanguageIDs(188) = msoLanguageIDSwissItalian
    LanguageIDs(189) = msoLanguageIDSyriac
    LanguageIDs(190) = msoLanguageIDTajik
    LanguageIDs(191) = msoLanguageIDTamazight
    LanguageIDs(192) = msoLanguageIDTamazightLatin
    LanguageIDs(193) = msoLanguageIDTamil
    LanguageIDs(194) = msoLanguageIDTatar
    LanguageIDs(195) = msoLanguageIDTelugu
    LanguageIDs(196) = msoLanguageIDThai
    LanguageIDs(197) = msoLanguageIDTibetan
    LanguageIDs(198) = msoLanguageIDTigrignaEritrea
    LanguageIDs(199) = msoLanguageIDTigrignaEthiopic
    LanguageIDs(200) = msoLanguageIDTraditionalChinese
    LanguageIDs(201) = msoLanguageIDTsonga
    LanguageIDs(202) = msoLanguageIDTswana
    LanguageIDs(203) = msoLanguageIDTurkish
    LanguageIDs(204) = msoLanguageIDTurkmen
    LanguageIDs(205) = msoLanguageIDUkrainian
    LanguageIDs(206) = msoLanguageIDUrdu
    LanguageIDs(207) = msoLanguageIDUzbekCyrillic
    LanguageIDs(208) = msoLanguageIDUzbekLatin
    LanguageIDs(209) = msoLanguageIDVenda
    LanguageIDs(210) = msoLanguageIDVietnamese
    LanguageIDs(211) = msoLanguageIDWelsh
    LanguageIDs(212) = msoLanguageIDXhosa
    LanguageIDs(213) = msoLanguageIDYi
    LanguageIDs(214) = msoLanguageIDYiddish
    LanguageIDs(215) = msoLanguageIDYoruba
    LanguageIDs(216) = msoLanguageIDZulu
    
    'Hide form
    ChangeSpellCheckLanguageForm.Hide
    
    Dim TargetLanguageID As String
    TargetLanguageID = LanguageIDs(ChangeSpellCheckLanguageForm.ComboBox1.ListIndex + 1)
    
    Dim TargetLanguage As String
    TargetLanguage = LanguageNames(ChangeSpellCheckLanguageForm.ComboBox1.ListIndex + 1)
    
    Dim PresentationSlide As PowerPoint.Slide
    Dim SlideShape  As PowerPoint.Shape
    Dim SlideSmartArtNode As SmartArtNode
    Dim GroupCount  As Integer
    
    ' Updates shapes in master
    For Each SlideShape In ActivePresentation.SlideMaster.Shapes
        If SlideShape.HasTextFrame Then
        SlideShape.TextFrame2.TextRange.LanguageID = TargetLanguageID
        End If
    Next
    
    For Each SlideShape In ActivePresentation.TitleMaster.Shapes
        If SlideShape.HasTextFrame Then
        SlideShape.TextFrame2.TextRange.LanguageID = TargetLanguageID
        End If
    Next
    
    For Each SlideShape In ActivePresentation.NotesMaster.Shapes
        If SlideShape.HasTextFrame Then
        SlideShape.TextFrame2.TextRange.LanguageID = TargetLanguageID
        End If
    Next
    
    ' Update shapes in slides
    For Each PresentationSlide In ActivePresentation.Slides
        
        For Each SlideShape In PresentationSlide.Shapes
            
            'Normal shapes
            If SlideShape.HasTextFrame Then
                SlideShape.TextFrame2.TextRange.LanguageID = TargetLanguageID
            End If
            
            'Tables
            If SlideShape.HasTable Then
                For TableRow = 1 To SlideShape.Table.Rows.Count
                    For TableColumn = 1 To SlideShape.Table.Columns.Count
                        SlideShape.Table.Cell(TableRow, TableColumn).Shape.TextFrame2.TextRange.LanguageID = TargetLanguageID
                    Next
                Next
                
                'SmartArt
                If SlideShape.HasSmartArt Then
                    For Each SlideSmartArtNode In SlideShape.SmartArt.AllNodes
                    If SlideSmartArdNode.HasTextFrame Then
                        SlideSmartArtNode.TextFrame2.TextRange.LanguageID = TargetLanguageID
                    End If
                    Next
                    
                End If
                
                'Groups - Note: need to find a better way to find out if it's a group.
                On Error Resume Next
                For GroupCount = 0 To SlideShape.GroupItems.Count - 1
                    SlideShape.GroupItems(GroupCount).TextFrame2.TextRange.LanguageID = TargetLanguageID
                Next
                
            End If
            
        Next SlideShape
        
    Next PresentationSlide
    
    MsgBox "Changed spellcheck language to " + TargetLanguage + " on all slides."
    
End Sub
