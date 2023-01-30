#pragma once

#include <iostream>
#include <fstream>
#include <sapi.h>
#pragma warning( push )
#pragma warning (disable:4996) //Aoife.Stroughair - Needed to disable a deprecation warning inside sphelper.h
#include <sphelper.h>
#pragma warning( pop )
#include <string>
#include <sstream>
#include <vector>
#include<direct.h>

#include "Public/TTSGenerator.h"


std::wstring StringToWString(const std::string sinputString)
{
	std::wstring ReturnValue(sinputString.size(), L' ');
	size_t outSize;
	ReturnValue.resize(mbstowcs_s(&outSize, &ReturnValue[0], sinputString.size() + 1, sinputString.c_str(), sinputString.size()));
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
			hr = TTSGenerator::GenerateTTS(TTSStrings[i]);
			if (!SUCCEEDED(hr))
			{
				std::cout << "An error occurred processing line" << i + 1 << std::endl;
			}
		}
		return 0;
	}
	else
	{
		HRESULT	hr = TTSGenerator::GenerateTTS(InputTTSString);
		if (!SUCCEEDED(hr))
		{
			std::cout << "An error occurred processing the inputted string\n";
			return hr;
		}
		return 0;
	}
}
