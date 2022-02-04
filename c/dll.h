/* dll.h */
#ifndef __DLL_H__
#define __DLL_H__

#define DLL_EXPORT __declspec(dllexport)
#include <Python.h>



     PyObject* DLL_EXPORT GetList(int abcLen, int MsgLen, int md); 
	 void DLL_EXPORT SetSeed();
	
#endif // __DLL_H__
