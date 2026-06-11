/*
 *	GPIB-Serial Tektronix scope screenshot tool
 *	By Tim Williams, 2021/05/11
 *	Updated 2024/09/13: minor poking at re-read function, sleep intervals.
 *	Updated 2026/06/10: HP8590A spectrum analyzer version.
 *
 *	Command line arguments:
 *		HPCAP [-p <port>] [-b <baud>] [-a <addr>] <output[.hgl]>
 *
 *		-p			Set port (default COM14)
 *		-b			Baud rate (default 230400);
 *					uses 8,N,1 serial configuration
 *		-a			Instrument GPIB address (default 18)
 *		<output>	Output file name.  If extension not given,
 *					.HGL is assumed (to write no extension, use '.')
 *					(TODO: autodetect by querying spec config)
 *
 *	Run with no parameters to see this help message.
 */

#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <tchar.h>
#include <shlwapi.h>


#define STRING_AND_LENGTH(x)		x, sizeof(x) - 1

#define MAX_TIMEOUTS		1
#define MS_START    		20
#define MS_WAIT     		20
#define MS_END      		200
#define MS_SHAKE    		1000

int main(int argc, char* argv[]);


HANDLE hComm;
HANDLE hOutput;
char portname[256] = "\\\\.\\COM14";
char filename[256];
char strBuf[1024];
long baud = 230400;
int address = 18;


int main(int argc, char* argv[]) {

	int i, err = 0;
	int numBytes;

	puts(
		"\n"
		"GPIB-Serial HP8590A spectrum analyzer plotting tool\n"
		"By Tim Williams, 2026/06/10\n"
		"\n"
	);

	if (argc <= 1) {

		puts(
			"Command line arguments:\n"
			"\tHPCAP [-p <port>] [-b <baud>] [-a <addr>] <output[.hgl]>\n"
			"\n"
			"    -p   \tSet port (default COM14)\n"
			"    -b   \tBaud rate (default 230400);\n"
			"         \tuses 8,N,1 serial configuration\n"
			"    -a   \tInstrument GPIB address (default 18)\n"
			" <output>\tOutput file name.  If extension not given, .HGL\n"
			"         \tis assumed (to write no extension, use '.').\n"
			"\n"
			"Run with no parameters to see this help message.\n"
		);

		return 0;
	}

	//	Iterate over parameters
	for (i = 1; i < argc - 1; i++) {

		if (!strcmp(argv[i], "-p")) {	//	Port name
			//	next token is the value, save it (keep the DOS device path specifier)
			i++;
			strncpy_s(portname + 4, sizeof(portname) - 1, argv[i], sizeof(portname) - 1);
		}

		if (!strcmp(argv[i], "-b")) {	//	Baud rate
			i++;
			baud = strtol(argv[i], NULL, 10);
		}

		if (!strcmp(argv[i], "-a")) {	//	Target GPIB address
			i++;
			address = strtol(argv[i], NULL, 10);
		}

	}

	//	Sanity checks
	if (i >= argc) {
		printf("Filename required.\n");
		return 1;
	}
	if (address > 30 || address < 0) {
		printf("Address %i out of range.\n", address);
		return 2;
	}
	if (baud < 0 || baud > 6000000l) {
		printf("Baud rate %li out of range.\n", baud);
		return 3;
	}
	//	Copy filename (it'll be parsed later, just before opening)
	strcpy_s(filename, sizeof(filename) - 1, argv[i]);

	//	Attempt to open the port
	hComm = CreateFile(portname,
				GENERIC_READ | GENERIC_WRITE,
				0,
				NULL,
				OPEN_EXISTING,
				0,
				NULL);

	if (hComm == INVALID_HANDLE_VALUE) {
		printf("Error opening port %s.\n", portname);
		err = 4; goto mainOut;
	}

	//	Configure port for settings
	DCB dcmComm = {
		.DCBlength			= sizeof(DCB),
		.BaudRate			= baud,
		.fBinary			= TRUE,
		.fParity			= FALSE,
		.fOutxCtsFlow		= FALSE,
		.fOutxDsrFlow		= FALSE,
		.fDtrControl		= FALSE,
		.fDsrSensitivity	= FALSE,
		.fTXContinueOnXoff	= FALSE,
		.fOutX				= FALSE,
		.fInX				= FALSE,
		.fErrorChar			= FALSE,
		.fNull				= FALSE,
		.fRtsControl		= RTS_CONTROL_DISABLE,
		.fAbortOnError		= FALSE,
		.XonLim				= FALSE,
		.XoffLim			= FALSE,
		.ByteSize			= 8,
		.Parity				= NOPARITY,
		.StopBits			= ONESTOPBIT,
		.XonChar			= 17,
		.XoffChar			= 19,
		.ErrorChar			= 255,
		.EofChar			= 3,
		.EvtChar			= 7,
	};
	err = SetCommState(hComm, &dcmComm);

	COMMTIMEOUTS ctComm = {
		.ReadIntervalTimeout			= 1000,
		.ReadTotalTimeoutMultiplier		= 0,
		.ReadTotalTimeoutConstant		= 1000,
		.WriteTotalTimeoutMultiplier	= 1,
		.WriteTotalTimeoutConstant		= 100
	};
	err = err && SetCommTimeouts(hComm, &ctComm);
	err = err && SetupComm(hComm, 1024, 1024);	//	Set minimum buffer sizes -- should be close to default anyway
	if (!err) {
		printf("IO error configuring port %s.\n", portname);
		err = 5; goto mainOut;
	}

	//	See if anyone's out there
	err = PurgeComm(hComm, PURGE_RXABORT | PURGE_RXCLEAR | PURGE_TXABORT | PURGE_TXCLEAR);
	err = err && WriteFile(hComm, STRING_AND_LENGTH("\r+read\r"), (LPDWORD)&numBytes, NULL);
	//	Make sure no pending transactions, clear buffer
	do {
		Sleep(MS_WAIT);
		if (!ReadFile(hComm, strBuf, sizeof(strBuf), (LPDWORD)&numBytes, NULL)) {
			printf("IO Error clearing input buffer.\n");
			err = 12; goto mainOut;
		}
	} while (numBytes > 0);
	err = err && WriteFile(hComm, STRING_AND_LENGTH("+ver\r"), (LPDWORD)&numBytes, NULL);
	Sleep(MS_START);
	err = err && ReadFile(hComm, strBuf, sizeof(strBuf), (LPDWORD)&numBytes, NULL);
	if (!err) {
		printf("IO error testing port.\n");
		err = 6; goto mainOut;
	}
	strBuf[numBytes] = 0;
	printf("GPIB adapter version: %s\n", strBuf);

	//	Append extension if none is found
	if (*PathFindExtensionA(filename) == 0) {
		strcat_s(filename, sizeof(filename) - 1, ".hgl");
	}

	//	Attempt to open output file
	hOutput = CreateFile(filename,
				GENERIC_WRITE,
				FILE_SHARE_READ,
				NULL,
				OPEN_ALWAYS,
				0,
				NULL);

	if (hOutput == INVALID_HANDLE_VALUE) {
		printf("IO error opening file %s.\n", filename);
		err = 7; goto mainOut;
	}

	//	we're going to assume `mode 1` is already set, since issuing it clears the instrument state
	sprintf_s(strBuf, sizeof(strBuf) - 1, "++addr %i\r++llo\r"/*"++mode 1\r"*/"PLOT 0 0 10250 7700;\r", address);
	if (!WriteFile(hComm, strBuf, strlen(strBuf), (LPDWORD)&numBytes, NULL)) {
		printf("IO error writing command.\n");
		err = 8; goto mainOut;
	}
	Sleep(MS_END);
	sprintf_s(strBuf, sizeof(strBuf) - 1, "+read\r", address);
	if (!WriteFile(hComm, strBuf, strlen(strBuf), (LPDWORD)&numBytes, NULL)) {
		printf("IO error writing command.\n");
		err = 8; goto mainOut;
	}

	long int lastTicks = GetTickCount();
	int bytesTilDot = 0;
	int timeout = 0;
	do {
		if (GetTickCount() - lastTicks > MS_SHAKE) {
			//	Seems to be taking a while to get more data... kick it with a +read?
			sprintf_s(strBuf, sizeof(strBuf) - 1, "+read\r");
			if (!WriteFile(hComm, strBuf, strlen(strBuf), (LPDWORD)&numBytes, NULL)) {
				printf("IO error during timeout retry.\n");
				err = 11; goto mainOut;
			}
			timeout++;
			lastTicks = GetTickCount();
			printf(":");
			Sleep(MS_WAIT);
		}
		if (!ReadFile(hComm, strBuf, sizeof(strBuf), (LPDWORD)&numBytes, NULL)) {
			printf("IO error reading data.\n");
			err = 9; goto mainOut;
		}
		if (numBytes > 0) {
			timeout = 0;
			if (!WriteFile(hOutput, strBuf, numBytes, (LPDWORD)&numBytes, NULL)) {
				printf("IO error writing output.\n");
				err = 10; goto mainOut;
			}
			bytesTilDot += numBytes;
			if (bytesTilDot > sizeof(strBuf)) {
				printf(".");
				bytesTilDot -= sizeof(strBuf);
			}
			lastTicks = GetTickCount();
		}
		Sleep(MS_WAIT);
	} while (timeout < MAX_TIMEOUTS);

	if (!WriteFile(hComm, STRING_AND_LENGTH("++loc\r++ifc\r++addr 0\r"), (LPDWORD)&numBytes, NULL)) {
		printf("IO error writing command.\n");
		err = 13; goto mainOut;
	}

	printf("\nDone.\n");
	Sleep(MS_END);

	err = 0;
mainOut:
	if (err) {
		//	Print the Windows error message and leave
		TCHAR *errStr;
		DWORD errCode = GetLastError();

		if (!FormatMessage(
					FORMAT_MESSAGE_ALLOCATE_BUFFER
					| FORMAT_MESSAGE_FROM_SYSTEM
					| FORMAT_MESSAGE_IGNORE_INSERTS,
					NULL,
					errCode,
					MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
					(LPTSTR)&errStr,
					0,
					NULL)) {
			printf("Unknown error 0x%lx occurred; extended error 0x%lx occurred getting error description.\n", errCode, GetLastError());
		} else {
			_tprintf_s("Error 0x%lx: %s", errCode, errStr);
		}
		LocalFree(errStr);
	}

	CloseHandle(hComm);
	CloseHandle(hOutput);
	Sleep(MS_END);

	return err;

}
