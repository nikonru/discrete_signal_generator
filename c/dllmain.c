/* Replace "dll.h" with the name of your header */
#include "dll.h"
#include <windows.h>
#include <Python.h>
#include <time.h>
#include <math.h>
//This .dll requires add to compiler options path to -lpython38 and Python.h
//Otherwise .dll won't be compiled

///
/*enum mode
{
	uniform,
	first_one_is_10_times_larger,
	first_seven_2_times_larger
};
*/


PyObject* DLL_EXPORT GetList(int abcLen, int MsgLen, int md)
{
	
	///
	int* letter_frequency;//массив с частотами появления каждого символа, для таблицы  

	float x_krit=-1;
	int n, k, t,msgLen;

	float H = 0;

	//int InitDSM(int MsgLen, int abcLen, enum mode md)


	letter_frequency = (int*)malloc(sizeof(int) * abcLen);
	//if (!letter_frequency)return -1;

	n = abcLen;
	msgLen = MsgLen;

	switch (md)
	{
	case 1:
		k = 0;
		t = 1;
		break;
	case 2:
		k = 1;
		t = 10;
		break;
	case 3:
		k = 7;
		t = 2;
		break;
	default:
		free(letter_frequency);
		//return -2;
		break;
	}


	x_krit = (float)(k*t)/(k*t+n-k);

	//int* GetMesseg() 
	float randomNum;

	int i;
	for (i = 0; i < n; i++)
		letter_frequency[i] = 0;

	for (i = 0; i < msgLen; i++)
	{
		randomNum = (float)(rand() % 1000) / 1000;
		if (x_krit > randomNum)
			letter_frequency[(int)(k * ((float)(rand() % 1000) / 1000))]++;
		else
			letter_frequency[k+(int)((n-k) * ((float)(rand() % 1000) / 1000))]++;
	}
	
	for (i = 0; i < n; i++){
		if(letter_frequency[i] !=0)
			H+= ((float)letter_frequency[i]/ msgLen)*log2f(((float)letter_frequency[i] / msgLen));
	}


	PyObject* py_list = PyList_New(abcLen+1);
	PyObject* value;

	for(i = 0; i < abcLen; i++)
	{
		value = Py_BuildValue("i",letter_frequency[i]);
		PyList_SetItem(py_list,i,value);
	}

	value = Py_BuildValue("f",-H);
	PyList_SetItem(py_list,abcLen,value);
	
	//void DeInitDSM()
	free(letter_frequency);
	
	return py_list;
}

void DLL_EXPORT SetSeed()
{
	srand(time(NULL));
}


BOOL WINAPI DllMain(HINSTANCE hinstDLL,DWORD fdwReason,LPVOID lpvReserved)
{
	switch(fdwReason)
	{
		case DLL_PROCESS_ATTACH:
		{
			break;
		}
		case DLL_PROCESS_DETACH:
		{
			break;
		}
		case DLL_THREAD_ATTACH:
		{
			break;
		}
		case DLL_THREAD_DETACH:
		{
			break;
		}
	}
	
	/* Return TRUE on success, FALSE on failure */
	return TRUE;
}
