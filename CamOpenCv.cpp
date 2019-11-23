
#include <time.h>

#define _WIN32_DCOM
#include <iostream>
#include <comdef.h>
#include <Wbemidl.h>
#pragma comment(lib, "wbemuuid.lib")
#include <opencv2/opencv.hpp>
int main(int argc, char* argv[])
{
	HRESULT hres;

	// Step 1: --------------------------------------------------
	// Initialize COM. ------------------------------------------

	hres = CoInitializeEx(0, COINIT_MULTITHREADED);
	if (FAILED(hres))
	{
		std::cout << "Failed to initialize COM library. Error code = 0x"
			<< std::hex << hres << std::endl;
		return 1;                  // Program has failed.
	}

	// Step 2: --------------------------------------------------
	// Set general COM security levels --------------------------

	hres = CoInitializeSecurity(
		NULL,
		-1,                          // COM negotiates service
		NULL,                        // Authentication services
		NULL,                        // Reserved
		RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
		RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation
		NULL,                        // Authentication info
		EOAC_NONE,                   // Additional capabilities 
		NULL                         // Reserved
	);


	if (FAILED(hres))
	{
		std::cout << "Failed to initialize security. Error code = 0x"
			<< std::hex << hres << std::endl;
		CoUninitialize();
		return 1;                      // Program has failed.
	}

	// Step 3: ---------------------------------------------------
	// Obtain the initial locator to WMI -------------------------

	IWbemLocator* pLoc = NULL;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		0,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID*)&pLoc);

	if (FAILED(hres))
	{
		std::cout << "Failed to create IWbemLocator object. "
			<< "Err code = 0x"
			<< std::hex << hres << std::endl;
		CoUninitialize();
		return 1;                 // Program has failed.
	}

	// Step 4: ---------------------------------------------------
	// Connect to WMI through the IWbemLocator::ConnectServer method

	IWbemServices* pSvc = NULL;

	// Connect to the local root\cimv2 namespace
	// and obtain pointer pSvc to make IWbemServices calls.
	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\WMI"),
		NULL,
		NULL,
		0,
		NULL,
		0,
		0,
		&pSvc
	);

	if (FAILED(hres))
	{
		std::cout << "Could not connect. Error code = 0x"
			<< std::hex << hres << std::endl;
		pLoc->Release();
		CoUninitialize();
		return 1;                // Program has failed.
	}

	std::cout << "Connected to ROOT\\WMI WMI namespace" << std::endl;


	// Step 5: --------------------------------------------------
	// Set security levels for the proxy ------------------------

	hres = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx 
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx 
		NULL,                        // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		NULL,                        // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hres))
	{
		std::cout << "Could not set proxy blanket. Error code = 0x"
			<< std::hex << hres << std::endl;
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return 1;               // Program has failed.
	}

	// Step 6: --------------------------------------------------
	// Use the IWbemServices pointer to make requests of WMI ----

	// set up to call the Win32_Process::Create method
	BSTR ClassName = SysAllocString(L"WmiMonitorBrightnessMethods");
	BSTR MethodName = SysAllocString(L"WmiSetBrightness");
	BSTR bstrQuery = SysAllocString(L"Select * from WmiMonitorBrightnessMethods");
	IEnumWbemClassObject* pEnum = NULL;

	hres = pSvc->ExecQuery(_bstr_t(L"WQL"), //Query Language  
		bstrQuery, //Query to Execute  
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, //Make a semi-synchronous call  
		NULL, //Context  
		&pEnum /*Enumeration Interface*/);

	hres = WBEM_S_NO_ERROR;

	ULONG ulReturned;
	IWbemClassObject* pObj;
	DWORD retVal = 0;

	//Get the Next Object from the collection  
	hres = pEnum->Next(WBEM_INFINITE, //Timeout  
		1, //No of objects requested  
		&pObj, //Returned Object  
		&ulReturned /*No of object returned*/);

	IWbemClassObject* pClass = NULL;
	hres = pSvc->GetObject(ClassName, 0, NULL, &pClass, NULL);

	IWbemClassObject* pInParamsDefinition = NULL;
	hres = pClass->GetMethod(MethodName, 0, &pInParamsDefinition, NULL);

	IWbemClassObject* pClassInstance = NULL;
	hres = pInParamsDefinition->SpawnInstance(0, &pClassInstance);

	VARIANT var1;
	VariantInit(&var1);
	BSTR ArgName0 = SysAllocString(L"Timeout");

	




	while (true) {
	cv::VideoCapture cap(0);
	cv::Mat frame;
	bool bSuccess = cap.read(frame);
	//Sleep(2000);
	if (cap.isOpened() == false)
	{
		std::cout << "Cannot open the video camera" << std::endl;
		std::cin.get(); //wait for any key press
		return -1;
	}

	std::string window_name = "My Camera Feed";
	//namedWindow(window_name); 

		
		int sum = 0;
		for (int i = 0; i >-1 ; i++)
		{
			
			//clock_t start = clock();
			Sleep(700);
			bSuccess = cap.read(frame); 
			//clock_t end = clock();
			int sredpix = 0;
			cv::cvtColor(frame, frame, cv::COLOR_BGR2GRAY);
			int kolpix = frame.rows*frame.cols/64 ;
			for (int k = 0; k < frame.rows;k+=8) {
				for (int l = 0; l < frame.cols; l += 8)
					sredpix += frame.at<cv::Vec3b>(k, l)[0];
			}
			std::cout << sredpix / kolpix << std::endl;
			V_VT(&var1) = VT_BSTR;

			V_BSTR(&var1) = SysAllocString(L"0");
			hres = pClassInstance->Put(ArgName0,
				0,
				&var1,
				CIM_UINT32); //CIM_UINT64  
			printf("\nPut ArgName0 returned 0x%x:", hres);
			VariantClear(&var1);

			VARIANT var2;
			VariantInit(&var2);
			BSTR ArgName1 = SysAllocString(L"Brightness");

			V_VT(&var2) = VT_BSTR;
			wchar_t temp_str[3];
			_itow_s(sredpix / kolpix/2, temp_str, 10);
			V_BSTR(&var2) = SysAllocString(temp_str); //Brightness value
			hres = pClassInstance->Put(ArgName1,
				0,
				&var2,
				CIM_UINT8);
			VariantClear(&var2);
			printf("\nPut ArgName1 returned 0x%x:", hres);

			// Call the method  
			VARIANT pathVariable;
			VariantInit(&pathVariable);

			hres = pObj->Get(_bstr_t(L"__PATH"),
				0,
				&pathVariable,
				NULL,
				NULL);
			printf("\npObj Get returned 0x%x:", hres);

			hres = pSvc->ExecMethod(pathVariable.bstrVal,
				MethodName,
				0,
				NULL,
				pClassInstance,
				NULL,
				NULL);
			VariantClear(&pathVariable);
			printf("\nExecMethod returned 0x%x:", hres);

			printf("Terminating normally\n");

			//if((end - start)<10
			//imshow(window_name, frame);
			//sum ++;
			if (bSuccess == false)
			{
				std::cout << "Video camera is disconnected" << std::endl;
				std::cin.get();
				break;
			}
			;

			
		}
		cap.release();
		
		//Sleep(20000);
		if (cv::waitKey(1) == 27)
		{
			std::cout << "Esc key is pressed by user. Stoppig the video" << std::endl;
			break;
		}
	}
	SysFreeString(ClassName);
	SysFreeString(MethodName);
	pClass->Release();
	pClassInstance->Release();
	pInParamsDefinition->Release();
	//pOutParams->Release();
	pLoc->Release();
	pSvc->Release();
	CoUninitialize();

	system("pause");
	return 0;

}