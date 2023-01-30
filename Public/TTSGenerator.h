#pragma once
#include <iostream>
#include <sapi.h>
#pragma warning( push )
#pragma warning (disable:4996) //Aoife.Stroughair - Needed to disable a deprecation warning inside sphelper.h
#include <sphelper.h>
#pragma warning( pop )
#include <string>

//Initial code originally taken from the Microsoft Speech API 5.4 whitepaper
//https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ee431810(v=vs.85)

#define OUTPUT_FOLDER L"OutputWAVs\\"

std::wstring ExePath() {
	TCHAR buffer[MAX_PATH] = { 0 };
	GetModuleFileName(NULL, buffer, MAX_PATH);
	std::wstring::size_type pos = std::wstring(buffer).find_last_of(L"\\/");
	return std::wstring(buffer).substr(0, pos);
}

struct TTSString
{
	std::wstring sInputString, sOutputFileName, sOutputPath, sVoiceAttributes;
};

class TTSGenerator
{
public:
	static int GenerateTTS(const TTSString& inputTTSString)
	{
		HRESULT					hr = S_OK;
		CComPtr <ISpVoice>		cpVoice;
		CComPtr <ISpStream>		cpStream;
		CSpStreamFormat			cAudioFmt;
		ISpObjectToken* cpToken(nullptr);

		std::wstring OutputFilePath;

		//Determine the output directory, and create it if necessary
		{
			if (!inputTTSString.sOutputPath.empty())
			{
				OutputFilePath = inputTTSString.sOutputPath + L'\\';
			}
			else
			{
				OutputFilePath = ExePath() + L'\\' + OUTPUT_FOLDER;
			}

			DWORD dwAttrib = GetFileAttributes(OutputFilePath.c_str());

			if (!(dwAttrib != INVALID_FILE_ATTRIBUTES && (dwAttrib & FILE_ATTRIBUTE_DIRECTORY)))
			{
				signed int pos = 0;
				do
				{
					pos = OutputFilePath.find_first_of(L"\\/", pos + 1);
					CreateDirectory(OutputFilePath.substr(0, pos).c_str(), NULL);
				} while (pos != std::wstring::npos);
			}

			OutputFilePath += inputTTSString.sOutputFileName;
		}

		hr = CoInitialize(nullptr);
		if (!SUCCEEDED(hr))
		{
			std::cout << "The call to CoInitialize() failed\n";
			return hr;
		}

		//Create a SAPI Voice
		hr = cpVoice.CoCreateInstance(CLSID_SpVoice);
		if (!SUCCEEDED(hr))
		{
			std::cout << "Failed to create a voice instance\n";
			return hr;
		}

		std::wstring sTokenAttribs = L"vendor=microsoft;";
		sTokenAttribs += inputTTSString.sVoiceAttributes;
		hr = SpFindBestToken(SPCAT_VOICES, sTokenAttribs.c_str(), L"", &cpToken);
		if (!SUCCEEDED(hr))
		{
			std::cout << "Failed to find a voice that worked with all the attributes listed. Removing excess attributes to find a default....\n";
			std::wstring sTokenAttribs = L"vendor=microsoft;";
			hr = SpFindBestToken(SPCAT_VOICES, sTokenAttribs.c_str(), L"", &cpToken);
			if (!SUCCEEDED(hr))
			{
				std::cout << "Failed to find a TTS voice\n";
				return hr;
			}
		}

		hr = cpVoice->SetVoice(cpToken);
		cpToken->Release();
		if (!SUCCEEDED(hr))
		{
			std::cout << "Failed to set the TTS voice\n";
			return hr;
		}

		//Set the audio format
		hr = cAudioFmt.AssignFormat(SPSF_22kHz16BitMono);
		if (!SUCCEEDED(hr))
		{
			std::cout << "Failed to set the audio format\n";
			return hr;
		}

		//Call SPBindToFile, a SAPI helper method,  to bind the audio stream to the file
		hr = SPBindToFile(OutputFilePath.c_str(), SPFM_CREATE_ALWAYS, &cpStream, &cAudioFmt.FormatId(), cAudioFmt.WaveFormatExPtr());
		if (!SUCCEEDED(hr))
		{
			std::cout << "Failed to bind the audio stream to the filepath (can files be written to this location?)\n";
			return hr;
		}

		//set the output to cpStream so that the output audio data will be stored in cpStream
		hr = cpVoice->SetOutput(cpStream, TRUE);
		if (!SUCCEEDED(hr))
		{
			std::cout << "Failed to set the output to the cpStream\n";
			return hr;
		}


		hr = cpVoice->Speak(inputTTSString.sInputString.c_str(), SPF_DEFAULT, NULL);
		if (!SUCCEEDED(hr))
		{
			std::cout << "Failed to begin writing the input string to the audio buffer\n";
			return hr;
		}

		hr = cpVoice->WaitUntilDone(1000);
		if (!SUCCEEDED(hr))
		{
			std::cout << "Failed to write the input string to the audio buffer\n";
			return hr;
		}

		//close the stream
		hr = cpStream->Close();
		if (!SUCCEEDED(hr))
		{
			std::cout << "Failed to close the audio stream\n";
			return hr;
		}

		//Release the stream and voice object
		cpStream.Release();
		cpVoice.Release();

		return hr;
	}
};