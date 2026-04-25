#define _CRT_SECURE_NO_WARNINGS
#define _WINSOCK_DEPRECATED_NO_WARNINGS
#define WIN32_LEAN_AND_MEAN
#define NTDDI_VERSION NTDDI_WIN10
#define _WIN32_WINNT _WIN32_WINNT_WIN10

#pragma warning(disable: 4996)

#include <iostream>
#include <string>
#include <vector>
#include <map>
#include <sstream>
#include <fstream>
#include <winsock2.h>
#include <ws2tcpip.h>
#include <sapi.h>
#include <sphelper.h>
#include <thread>
#include <mutex>

#pragma comment(lib, "ws2_32.lib")
#pragma comment(lib, "sapi.lib")

#define MAX_BUFFER 8192
#define PORT 8000

typedef struct lame_global_flags* (*lame_init_t)(void);
typedef int (*lame_set_num_channels_t)(struct lame_global_flags*, int);
typedef int (*lame_set_in_samplerate_t)(struct lame_global_flags*, int);
typedef int (*lame_set_mode_t)(struct lame_global_flags*, int);
typedef int (*lame_init_params_t)(struct lame_global_flags*);
typedef int (*lame_encode_buffer_t)(struct lame_global_flags*, const short int*, const short int*, int, unsigned char*, int);
typedef int (*lame_encode_flush_t)(struct lame_global_flags*, unsigned char*, int);
typedef void (*lame_close_t)(struct lame_global_flags*);

static lame_init_t lame_init = NULL;
static lame_set_num_channels_t lame_set_num_channels = NULL;
static lame_set_in_samplerate_t lame_set_in_samplerate = NULL;
static lame_set_mode_t lame_set_mode = NULL;
static lame_init_params_t lame_init_params = NULL;
static lame_encode_buffer_t lame_encode_buffer = NULL;
static lame_encode_flush_t lame_encode_flush = NULL;
static lame_close_t lame_close = NULL;
static HMODULE hLameDll = NULL;

bool LoadLameLibrary() {
    hLameDll = LoadLibraryA("lame_enc.dll");
    if (!hLameDll) {
        hLameDll = LoadLibraryA("libmp3lame.dll");
    }
    if (!hLameDll) {
        std::cerr << "Cannot load lame_enc.dll or libmp3lame.dll" << std::endl;
        return false;
    }
    
    lame_init = (lame_init_t)GetProcAddress(hLameDll, "lame_init");
    lame_set_num_channels = (lame_set_num_channels_t)GetProcAddress(hLameDll, "lame_set_num_channels");
    lame_set_in_samplerate = (lame_set_in_samplerate_t)GetProcAddress(hLameDll, "lame_set_in_samplerate");
    lame_set_mode = (lame_set_mode_t)GetProcAddress(hLameDll, "lame_set_mode");
    lame_init_params = (lame_init_params_t)GetProcAddress(hLameDll, "lame_init_params");
    lame_encode_buffer = (lame_encode_buffer_t)GetProcAddress(hLameDll, "lame_encode_buffer");
    lame_encode_flush = (lame_encode_flush_t)GetProcAddress(hLameDll, "lame_encode_flush");
    lame_close = (lame_close_t)GetProcAddress(hLameDll, "lame_close");
    
    if (!lame_init || !lame_set_num_channels || !lame_set_in_samplerate ||
        !lame_init_params || !lame_encode_buffer || !lame_encode_flush || !lame_close) {
        std::cerr << "Cannot get lame function addresses" << std::endl;
        FreeLibrary(hLameDll);
        hLameDll = NULL;
        return false;
    }
    
    return true;
}

void UnloadLameLibrary() {
    if (hLameDll) {
        FreeLibrary(hLameDll);
        hLameDll = NULL;
    }
}

std::wstring Utf8ToWide(const std::string& utf8) {
    if (utf8.empty()) return L"";
    int size = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, NULL, 0);
    std::wstring result(size - 1, 0);
    MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, &result[0], size);
    return result;
}

std::string WideToUtf8(const std::wstring& wide) {
    if (wide.empty()) return "";
    int size = WideCharToMultiByte(CP_UTF8, 0, wide.c_str(), -1, NULL, 0, NULL, NULL);
    std::string result(size - 1, 0);
    WideCharToMultiByte(CP_UTF8, 0, wide.c_str(), -1, &result[0], size, NULL, NULL);
    return result;
}

std::string UrlDecode(const std::string& encoded) {
    std::string decoded;
    char ch;
    int i, ii;
    for (i = 0; i < (int)encoded.length(); i++) {
        if (encoded[i] == '%') {
            if (i + 2 >= (int)encoded.length()) break;
            std::string hex = encoded.substr(i + 1, 2);
            ii = (int)strtol(hex.c_str(), NULL, 16);
            ch = static_cast<char>(ii);
            decoded += ch;
            i += 2;
        } else if (encoded[i] == '+') {
            decoded += ' ';
        } else {
            decoded += encoded[i];
        }
    }
    return decoded;
}

std::vector<std::pair<std::wstring, std::wstring>> GetVoiceList() {
    std::vector<std::pair<std::wstring, std::wstring>> voices;
    ISpVoice* pVoice = NULL;
    
    if (FAILED(CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice))) {
        return voices;
    }
    
    ISpObjectTokenCategory* pCategory = NULL;
    if (SUCCEEDED(CoCreateInstance(CLSID_SpObjectTokenCategory, NULL, CLSCTX_ALL, IID_ISpObjectTokenCategory, (void**)&pCategory))) {
        pCategory->SetId(SPCAT_VOICES, FALSE);
        
        IEnumSpObjectTokens* pEnum = NULL;
        if (SUCCEEDED(pCategory->EnumTokens(NULL, NULL, &pEnum))) {
            ISpObjectToken* pToken = NULL;
            ULONG ulFetched = 0;
            
            while (SUCCEEDED(pEnum->Next(1, &pToken, &ulFetched)) && ulFetched == 1) {
                WCHAR* pId = NULL;
                if (SUCCEEDED(pToken->GetId(&pId))) {
                    std::wstring id = pId;
                    CoTaskMemFree(pId);
                    
                    WCHAR* pDesc = NULL;
                    if (SUCCEEDED(SpGetDescription(pToken, &pDesc))) {
                        std::wstring desc = pDesc;
                        CoTaskMemFree(pDesc);
                        voices.push_back(std::make_pair(id, desc));
                    }
                }
                pToken->Release();
            }
            pEnum->Release();
        }
        pCategory->Release();
    }
    
    pVoice->Release();
    return voices;
}

std::vector<BYTE> TextToWav(const std::wstring& text, const std::wstring& voiceId) {
    std::vector<BYTE> wavData;
    
    ISpVoice* pVoice = NULL;
    if (FAILED(CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice))) {
        return wavData;
    }
    
    if (!voiceId.empty()) {
        ISpObjectToken* pToken = NULL;
        if (SUCCEEDED(SpGetTokenFromId(voiceId.c_str(), &pToken, FALSE))) {
            pVoice->SetVoice(pToken);
            pToken->Release();
        }
    }
    
    CSpStreamFormat format;
    format.AssignFormat(SPSF_22kHz16BitMono);
    
    HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, 0);
    if (!hGlobal) {
        pVoice->Release();
        return wavData;
    }
    
    IStream* pMemStream = NULL;
    if (FAILED(CreateStreamOnHGlobal(hGlobal, TRUE, &pMemStream))) {
        GlobalFree(hGlobal);
        pVoice->Release();
        return wavData;
    }
    
    ISpStream* pSpStream = NULL;
    if (FAILED(CoCreateInstance(CLSID_SpStream, NULL, CLSCTX_ALL, IID_ISpStream, (void**)&pSpStream))) {
        pMemStream->Release();
        pVoice->Release();
        return wavData;
    }
    
    if (FAILED(pSpStream->SetBaseStream(pMemStream, SPDFID_WaveFormatEx, format.WaveFormatExPtr()))) {
        pSpStream->Release();
        pMemStream->Release();
        pVoice->Release();
        return wavData;
    }
    
    pVoice->SetOutput(pSpStream, TRUE);
    
    if (SUCCEEDED(pVoice->Speak(text.c_str(), SPF_DEFAULT, NULL))) {
        pVoice->WaitUntilDone(INFINITE);
        pVoice->SetOutput(NULL, FALSE);
        
        LARGE_INTEGER liZero = {0};
        ULARGE_INTEGER ulPos;
        pMemStream->Seek(liZero, STREAM_SEEK_CUR, &ulPos);
        DWORD dwSize = (DWORD)ulPos.QuadPart;
        
        pMemStream->Seek(liZero, STREAM_SEEK_SET, NULL);
        
        wavData.resize(dwSize);
        ULONG ulRead;
        pMemStream->Read(wavData.data(), dwSize, &ulRead);
    } else {
        pVoice->SetOutput(NULL, FALSE);
    }
    
    pSpStream->Release();
    pMemStream->Release();
    pVoice->Release();
    
    return wavData;
}

std::vector<BYTE> ConvertWavToMp3(const std::vector<BYTE>& wavData) {
    std::vector<BYTE> mp3Data;
    
    if (!hLameDll || !lame_init) {
        std::cerr << "LAME library not loaded, cannot convert to MP3" << std::endl;
        return mp3Data;
    }
    
    if (wavData.size() < 44) {
        return mp3Data;
    }
    
    const BYTE* wavHeader = wavData.data();
    int sampleRate = *(int*)(wavHeader + 24);
    int numChannels = wavHeader[22];
    int bitsPerSample = wavHeader[34];
    
    if (bitsPerSample != 16) {
        std::cerr << "Only 16-bit WAV format supported" << std::endl;
        return mp3Data;
    }
    
    int dataOffset = 44;
    for (int i = 12; i < (int)wavData.size() - 8; i++) {
        if (wavData[i] == 'd' && wavData[i+1] == 'a' && wavData[i+2] == 't' && wavData[i+3] == 'a') {
            dataOffset = i + 8;
            break;
        }
    }
    
    const short* pcmData = (const short*)(wavData.data() + dataOffset);
    int pcmSamples = (int)((wavData.size() - dataOffset) / 2 / numChannels);
    
    auto gf = lame_init();
    if (!gf) {
        return mp3Data;
    }
    
    lame_set_num_channels(gf, numChannels);
    lame_set_in_samplerate(gf, sampleRate);
    if (lame_set_mode) {
        lame_set_mode(gf, numChannels == 1 ? 3 : 1);
    }
    
    if (lame_init_params(gf) < 0) {
        lame_close(gf);
        return mp3Data;
    }
    
    int mp3BufferSize = (int)(pcmSamples * 1.25 + 7200);
    mp3Data.resize(mp3BufferSize);
    
    int mp3Size;
    if (numChannels == 1) {
        mp3Size = lame_encode_buffer(gf, pcmData, NULL, pcmSamples, mp3Data.data(), mp3BufferSize);
    } else {
        mp3Size = lame_encode_buffer(gf, pcmData, pcmData + 1, pcmSamples, mp3Data.data(), mp3BufferSize);
    }
    
    if (mp3Size < 0) {
        lame_close(gf);
        mp3Data.clear();
        return mp3Data;
    }
    
    int flushSize = lame_encode_flush(gf, mp3Data.data() + mp3Size, mp3BufferSize - mp3Size);
    if (flushSize > 0) {
        mp3Size += flushSize;
    }
    
    lame_close(gf);
    
    mp3Data.resize(mp3Size);
    return mp3Data;
}

std::string GetHttpResponse(const std::string& content, const std::string& contentType, int statusCode = 200) {
    std::ostringstream response;
    response << "HTTP/1.1 " << statusCode << " OK\r\n";
    response << "Content-Type: " << contentType << "\r\n";
    response << "Content-Length: " << content.length() << "\r\n";
    response << "Connection: close\r\n";
    response << "Access-Control-Allow-Origin: *\r\n";
    response << "\r\n";
    response << content;
    return response.str();
}

std::string GetBinaryHttpResponse(const std::vector<BYTE>& data, const std::string& contentType, const std::string& disposition = "") {
    std::ostringstream headers;
    headers << "HTTP/1.1 200 OK\r\n";
    headers << "Content-Type: " << contentType << "\r\n";
    headers << "Content-Length: " << data.size() << "\r\n";
    headers << "Connection: close\r\n";
    headers << "Access-Control-Allow-Origin: *\r\n";
    if (!disposition.empty()) {
        headers << "Content-Disposition: " << disposition << "\r\n";
    }
    headers << "\r\n";
    
    std::string headersStr = headers.str();
    std::string response;
    response.reserve(headersStr.size() + data.size());
    response.append(headersStr);
    response.append((const char*)data.data(), data.size());
    return response;
}

std::string GetHtmlPage() {
    return R"(
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Text to Speech</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Microsoft YaHei', Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 16px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            padding: 40px;
            width: 100%;
            max-width: 600px;
        }
        h1 {
            text-align: center;
            color: #333;
            margin-bottom: 30px;
            font-size: 28px;
        }
        .form-group {
            margin-bottom: 20px;
        }
        label {
            display: block;
            margin-bottom: 8px;
            color: #555;
            font-weight: 600;
        }
        textarea {
            width: 100%;
            height: 150px;
            padding: 15px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            font-size: 16px;
            resize: vertical;
            transition: border-color 0.3s;
            font-family: inherit;
        }
        textarea:focus {
            outline: none;
            border-color: #667eea;
        }
        select {
            width: 100%;
            padding: 12px 15px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            font-size: 16px;
            background: white;
            cursor: pointer;
            transition: border-color 0.3s;
        }
        select:focus {
            outline: none;
            border-color: #667eea;
        }
        .btn-group {
            display: flex;
            gap: 15px;
            margin-top: 30px;
        }
        button {
            flex: 1;
            padding: 15px;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }
        .btn-play {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        .btn-play:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 20px rgba(102, 126, 234, 0.4);
        }
        .btn-download {
            background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
            color: white;
        }
        .btn-download:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 20px rgba(17, 153, 142, 0.4);
        }
        button:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
        .status {
            margin-top: 20px;
            padding: 15px;
            border-radius: 8px;
            text-align: center;
            display: none;
        }
        .status.loading {
            background: #e3f2fd;
            color: #1976d2;
            display: block;
        }
        .status.error {
            background: #ffebee;
            color: #c62828;
            display: block;
        }
        .status.success {
            background: #e8f5e9;
            color: #2e7d32;
            display: block;
        }
        .audio-container {
            margin-top: 20px;
            display: none;
        }
        audio {
            width: 100%;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Text to Speech</h1>
        
        <div class="form-group">
            <label for="voice">Select Voice:</label>
            <select id="voice">
                <option value="">Loading...</option>
            </select>
        </div>
        
        <div class="form-group">
            <label for="text">Enter Text:</label>
            <textarea id="text" placeholder="Enter text to convert..."></textarea>
        </div>
        
        <div class="btn-group">
            <button class="btn-play" id="btnPlay">
                <span>Play</span> Convert
            </button>
            <button class="btn-download" id="btnDownload">
                <span>Download</span> MP3
            </button>
        </div>
        
        <div class="status" id="status"></div>
        
        <div class="audio-container" id="audioContainer">
            <audio id="audioPlayer" controls></audio>
        </div>
    </div>
    
    <script>
        const textarea = document.getElementById('text');
        const voiceSelect = document.getElementById('voice');
        const btnPlay = document.getElementById('btnPlay');
        const btnDownload = document.getElementById('btnDownload');
        const statusDiv = document.getElementById('status');
        const audioContainer = document.getElementById('audioContainer');
        const audioPlayer = document.getElementById('audioPlayer');
        
        let currentAudioUrl = null;
        
        function showStatus(message, type) {
            statusDiv.textContent = message;
            statusDiv.className = 'status ' + type;
        }
        
        function hideStatus() {
            statusDiv.className = 'status';
        }
        
        function setButtonsEnabled(enabled) {
            btnPlay.disabled = !enabled;
            btnDownload.disabled = !enabled;
        }
        
        async function loadVoices() {
            try {
                const response = await fetch('/voices');
                const voices = await response.json();
                
                voiceSelect.innerHTML = '';
                voices.forEach(voice => {
                    const option = document.createElement('option');
                    option.value = voice.id;
                    option.textContent = voice.name;
                    voiceSelect.appendChild(option);
                });
                
                if (voices.length === 0) {
                    const option = document.createElement('option');
                    option.value = '';
                    option.textContent = 'No voices available';
                    voiceSelect.appendChild(option);
                }
            } catch (error) {
                console.error('Failed to load voices:', error);
                voiceSelect.innerHTML = '<option value="">Load failed</option>';
            }
        }
        
        async function convertText(format) {
            const text = textarea.value.trim();
            if (!text) {
                showStatus('Please enter text', 'error');
                return;
            }
            
            const voiceId = voiceSelect.value;
            setButtonsEnabled(false);
            hideStatus();
            
            try {
                showStatus('Converting...', 'loading');
                
                const formData = new URLSearchParams();
                formData.append('text', text);
                if (voiceId) {
                    formData.append('voice', voiceId);
                }
                
                const endpoint = format === 'wav' ? '/speak' : '/download';
                const response = await fetch(endpoint, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded'
                    },
                    body: formData
                });
                
                if (!response.ok) {
                    throw new Error('Conversion failed');
                }
                
                const blob = await response.blob();
                
                if (currentAudioUrl) {
                    URL.revokeObjectURL(currentAudioUrl);
                }
                
                currentAudioUrl = URL.createObjectURL(blob);
                
                if (format === 'wav') {
                    audioPlayer.src = currentAudioUrl;
                    audioContainer.style.display = 'block';
                    audioPlayer.play();
                    showStatus('Conversion successful!', 'success');
                } else {
                    const a = document.createElement('a');
                    a.href = currentAudioUrl;
                    a.download = 'tts_' + Date.now() + '.mp3';
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    showStatus('Download started!', 'success');
                }
                
            } catch (error) {
                showStatus('Conversion failed: ' + error.message, 'error');
            } finally {
                setButtonsEnabled(true);
            }
        }
        
        btnPlay.addEventListener('click', () => convertText('wav'));
        btnDownload.addEventListener('click', () => convertText('mp3'));
        
        loadVoices();
    </script>
</body>
</html>
)";
}

std::string ParseRequest(const std::string& request, std::string& method, std::string& path, std::map<std::string, std::string>& headers, std::string& body) {
    std::istringstream stream(request);
    std::string line;
    
    if (std::getline(stream, line)) {
        std::istringstream lineStream(line);
        lineStream >> method >> path;
    }
    
    while (std::getline(stream, line) && line != "\r" && line != "") {
        size_t colon = line.find(':');
        if (colon != std::string::npos) {
            std::string key = line.substr(0, colon);
            std::string value = line.substr(colon + 1);
            
            size_t start = value.find_first_not_of(" \t");
            if (start != std::string::npos) {
                value = value.substr(start);
            }
            size_t end = value.find_last_not_of(" \t\r");
            if (end != std::string::npos) {
                value = value.substr(0, end + 1);
            }
            
            headers[key] = value;
        }
    }
    
    auto contentLengthIt = headers.find("Content-Length");
    if (contentLengthIt != headers.end()) {
        int length = atoi(contentLengthIt->second.c_str());
        if (length > 0) {
            char* buffer = new char[length + 1];
            stream.read(buffer, length);
            buffer[length] = '\0';
            body = buffer;
            delete[] buffer;
        }
    }
    
    return path;
}

std::map<std::string, std::string> ParseFormData(const std::string& body) {
    std::map<std::string, std::string> data;
    std::istringstream stream(body);
    std::string pair;
    
    while (std::getline(stream, pair, '&')) {
        size_t eq = pair.find('=');
        if (eq != std::string::npos) {
            std::string key = pair.substr(0, eq);
            std::string value = pair.substr(eq + 1);
            data[key] = UrlDecode(value);
        }
    }
    
    return data;
}

void HandleClient(SOCKET clientSocket) {
    char buffer[MAX_BUFFER];
    std::string request;
    int bytesReceived;
    
    while ((bytesReceived = recv(clientSocket, buffer, MAX_BUFFER - 1, 0)) > 0) {
        buffer[bytesReceived] = '\0';
        request += buffer;
        
        if (request.find("\r\n\r\n") != std::string::npos) {
            break;
        }
    }
    
    if (bytesReceived <= 0 && request.empty()) {
        closesocket(clientSocket);
        return;
    }
    
    std::string method, path, body;
    std::map<std::string, std::string> headers;
    ParseRequest(request, method, path, headers, body);
    
    std::cout << "[" << method << "] " << path << std::endl;
    
    std::string response;
    
    if (method == "GET" && (path == "/" || path == "/index.html")) {
        std::string html = GetHtmlPage();
        response = GetHttpResponse(html, "text/html; charset=utf-8");
    }
    else if (method == "GET" && path == "/voices") {
        auto voices = GetVoiceList();
        std::ostringstream json;
        json << "[";
        for (size_t i = 0; i < voices.size(); i++) {
            if (i > 0) json << ",";
            json << "{\"id\":\"" << WideToUtf8(voices[i].first) << "\",";
            json << "\"name\":\"" << WideToUtf8(voices[i].second) << "\"}";
        }
        json << "]";
        response = GetHttpResponse(json.str(), "application/json; charset=utf-8");
    }
    else if (method == "POST" && (path == "/speak" || path == "/download")) {
        auto formData = ParseFormData(body);
        auto textIt = formData.find("text");
        
        if (textIt == formData.end() || textIt->second.empty()) {
            response = GetHttpResponse("Please enter text", "text/plain", 400);
        } else {
            std::wstring text = Utf8ToWide(textIt->second);
            std::wstring voiceId;
            auto voiceIt = formData.find("voice");
            if (voiceIt != formData.end()) {
                voiceId = Utf8ToWide(voiceIt->second);
            }
            
            std::vector<BYTE> audioData = TextToWav(text, voiceId);
            
            if (audioData.empty()) {
                response = GetHttpResponse("Speech conversion failed", "text/plain", 500);
            } else {
                if (path == "/download") {
                    std::vector<BYTE> mp3Data = ConvertWavToMp3(audioData);
                    if (!mp3Data.empty()) {
                        response = GetBinaryHttpResponse(mp3Data, "audio/mpeg", "attachment; filename=tts.mp3");
                    } else {
                        response = GetBinaryHttpResponse(audioData, "audio/wav", "attachment; filename=tts.wav");
                    }
                } else {
                    response = GetBinaryHttpResponse(audioData, "audio/wav");
                }
            }
        }
    }
    else {
        response = GetHttpResponse("404 Not Found", "text/plain", 404);
    }
    
    send(clientSocket, response.c_str(), (int)response.length(), 0);
    closesocket(clientSocket);
}

int main() {
    SetConsoleOutputCP(CP_UTF8);
    
    std::cout << "Initializing COM..." << std::endl;
    CoInitialize(NULL);
    
    std::cout << "Loading LAME encoder..." << std::endl;
    LoadLameLibrary();
    if (!hLameDll) {
        std::cout << "Warning: lame_enc.dll not found, download will use WAV format" << std::endl;
        std::cout << "For MP3 support, place lame_enc.dll or libmp3lame.dll in the same directory" << std::endl;
    } else {
        std::cout << "LAME encoder loaded successfully, MP3 format supported" << std::endl;
    }
    
    WSADATA wsaData;
    if (WSAStartup(MAKEWORD(2, 2), &wsaData) != 0) {
        std::cerr << "WSAStartup failed" << std::endl;
        CoUninitialize();
        return 1;
    }
    
    SOCKET serverSocket = socket(AF_INET, SOCK_STREAM, 0);
    if (serverSocket == INVALID_SOCKET) {
        std::cerr << "socket creation failed" << std::endl;
        WSACleanup();
        CoUninitialize();
        return 1;
    }
    
    sockaddr_in serverAddr;
    serverAddr.sin_family = AF_INET;
    serverAddr.sin_addr.s_addr = INADDR_ANY;
    serverAddr.sin_port = htons(PORT);
    
    if (bind(serverSocket, (sockaddr*)&serverAddr, sizeof(serverAddr)) == SOCKET_ERROR) {
        std::cerr << "bind failed" << std::endl;
        closesocket(serverSocket);
        WSACleanup();
        CoUninitialize();
        return 1;
    }
    
    if (listen(serverSocket, 10) == SOCKET_ERROR) {
        std::cerr << "listen failed" << std::endl;
        closesocket(serverSocket);
        WSACleanup();
        CoUninitialize();
        return 1;
    }
    
    std::cout << "========================================" << std::endl;
    std::cout << "Text-to-Speech Server Started" << std::endl;
    std::cout << "Listening on port: " << PORT << std::endl;
    std::cout << "Please visit: http://localhost:" << PORT << std::endl;
    std::cout << "========================================" << std::endl;
    
    while (true) {
        sockaddr_in clientAddr;
        int clientLen = sizeof(clientAddr);
        SOCKET clientSocket = accept(serverSocket, (sockaddr*)&clientAddr, &clientLen);
        
        if (clientSocket == INVALID_SOCKET) {
            std::cerr << "accept failed" << std::endl;
            continue;
        }
        
        char clientIP[INET_ADDRSTRLEN];
        inet_ntop(AF_INET, &clientAddr.sin_addr, clientIP, INET_ADDRSTRLEN);
        std::cout << "Client connected: " << clientIP << ":" << ntohs(clientAddr.sin_port) << std::endl;
        
        std::thread t(HandleClient, clientSocket);
        t.detach();
    }
    
    closesocket(serverSocket);
    WSACleanup();
    UnloadLameLibrary();
    CoUninitialize();
    
    return 0;
}
