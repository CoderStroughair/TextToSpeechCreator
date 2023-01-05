#pragma once

#include <iostream>
#include <fstream>
#include <sapi.h>
#pragma warning (disable:4996) //Aoife.Stroughair - Needed to disable a deprecation warning inside sphelper.h
#include <sphelper.h>
#include <string>
#include <sstream>
#include <vector>
#include<direct.h>

//Initial code originally taken from the Microsoft Speech API 5.4 whitepaper
//https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ee431810(v=vs.85)

#define OUTPUT_FOLDER L"OutputWAVs\\"

struct TTSString 
{
	std::wstring sInputString, sOutputFileName, sOutputPath, sVoiceAttributes;
};

std::wstring ExePath() {
	TCHAR buffer[MAX_PATH] = { 0 };
	GetModuleFileName(NULL, buffer, MAX_PATH);
	std::wstring::size_type pos = std::wstring(buffer).find_last_of(L"\\/");
	return std::wstring(buffer).substr(0, pos);
}

int GenerateTTS(const TTSString& inputTTSString)
{
	HRESULT					hr = S_OK;
	CComPtr <ISpVoice>		cpVoice;
	CComPtr <ISpStream>		cpStream;
	CSpStreamFormat			cAudioFmt;
	ISpObjectToken*			cpToken(nullptr);

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

std::wstring StringToWString(const std::string sinputString)
{
	std::wstring ReturnValue(sinputString.size(), L' ');
	ReturnValue.resize(std::mbstowcs(&ReturnValue[0], sinputString.c_str(), sinputString.size()));
	return ReturnValue;
}

HRESULT ImportCSV(const std::wstring& sOutputFileName, std::vector<TTSString>& TTSStrings)
{
	// Read from the text file
	std::wifstream MyReadFile(sOutputFileName.c_str());

	for (std::wstring CSVLine; std::getline(MyReadFile, CSVLine); ) 
	{
		std::wstringstream ss(CSVLine.c_str());
		std::wstring CSVItem;
		TTSString currTTSString;

		std::getline(ss, CSVItem, L';');
		currTTSString.sInputString = CSVItem;
		std::getline(ss, CSVItem, L';');
		currTTSString.sOutputFileName = CSVItem;
		std::getline(ss, CSVItem, L';');
		currTTSString.sOutputPath = CSVItem;
		std::getline(ss, CSVItem, L';');
		currTTSString.sVoiceAttributes = CSVItem;

		if (currTTSString.sInputString.empty() || currTTSString.sOutputFileName.empty())
		{
			std::cout << "The program has read in an incomplete CSV line. This will be discarded.\n";
			continue;
		}
		TTSStrings.push_back(currTTSString);
	}

	// Close the file
	MyReadFile.close();

	if (TTSStrings.size() == 0)
	{
		std::cout << "The program has not managed to read in any lines, please check your CSV and try again.\n";
		return ERROR_FILE_CORRUPT;
	}
	return S_OK;
}

int wmain(int argc, wchar_t* argv[])
{
	if ((argc == 1) || (argv[1] == L"-h"))
	{
		std::cout << "Usage Guide:\n \
		Ensure that all arguments use double quotes (\"\") around them so they are treated as a single argument\n\
		\n\
		\n\
		Parameters\n\
		-CSVFileName:\"<CSV file name>\" - Name of the CSV file you intend to read from. Passing in this parameter will cause all other parameters to be ignored.\n \
		-CSVFilePath:\"<CSV file path>\" - Path to the CSV file you intend to read from. If this is not passed the program will assume the file is in the same directory as the exe\n \
		-OutputFileName:\"<Output wav file name>\" - Name of the file you wish to output, incl. extension \n \
		-OutputFilePath:\"<Output wav file path>\" - The path to which you wish to save the output file \n \
		-InputString:\"<String to TTSify>\" - The string that you wish to convert to speech \n \
		-VoiceAttributes:\"<Voice Parameters>\" - parameters used by SAPI to locate the voice required.Vendor is automatically set to Microsoft \n \
		-h - Brings up this help menu\n\
		";
		return 0;
	}

	TTSString InputTTSString;
	std::wstring CSVFileName, CSVFilePath;

	for (int i = 1; i < argc; i++)
	{
		//first ensure that all parameters are valid
		if (argv[i][0] != '-')
		{
			std::cout << "Invalid parameters passed to the application. It will now terminate\n";
			return -1;
		}
		wchar_t* strNextToken = nullptr;
		const wchar_t* strToken = wcstok_s(argv[i]+1, L":", &strNextToken);
		if (strToken == 0 || strToken == L"")
		{
			std::cout << "Invalid parameters passed to the application. It will now terminate\n";
			return -1;
		}

		if (wcsstr(strToken, L"CSVFileName") == strToken)
		{
			CSVFileName = strNextToken;
		}
		else if (wcsstr(strToken, L"CSVFilePath") == strToken)
		{
			CSVFilePath = strNextToken;
		}
		else if (wcsstr(strToken, L"OutputFileName") == strToken)
		{
			InputTTSString.sOutputFileName = strNextToken;
		}
		else if (wcsstr(strToken, L"OutputFilePath") == strToken)
		{
			InputTTSString.sOutputPath = strNextToken;
		}
		else if (wcsstr(strToken, L"InputString") == strToken)
		{
			InputTTSString.sInputString = strNextToken;
		}
		else if (wcsstr(strToken, L"VoiceAttributes") == strToken)
		{
			InputTTSString.sVoiceAttributes = strNextToken;
		}
	}

	if (!CSVFileName.empty())
	{
		if (CSVFilePath.empty())
		{
			CSVFileName = ExePath() + L'\\' + CSVFileName;
		}
		else
		{
			CSVFileName = CSVFilePath + L'\\' + CSVFileName;
		}

		HRESULT	hr = S_OK;

		std::vector<TTSString> TTSStrings;
		hr = ImportCSV(CSVFileName, TTSStrings);
		if (!SUCCEEDED(hr))
		{
			std::cout << "An error occurred importing the CSV file.\n";
			return hr;
		}

		for (unsigned int i = 0; i < TTSStrings.size(); i++)
		{
			hr = GenerateTTS(TTSStrings[i]);
			if (!SUCCEEDED(hr))
			{
				std::cout << "An error occurred processing line" << i + 1 << std::endl;
			}
		}
		return 0;
	}
	else
	{
		HRESULT	hr = GenerateTTS(InputTTSString);
		if (!SUCCEEDED(hr))
		{
			std::cout << "An error occurred processing the inputted string\n";
			return hr;
		}
		return 0;
	}
}
