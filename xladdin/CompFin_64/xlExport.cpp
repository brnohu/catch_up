
#include <windows.h>
#include "xlcall.h"
#include "framework.h"
#include "xlOper.h"
#include "../Utility/kMatrixAlgebra.h"
#include "../Utility/xlUtils.h"

//	Wrappers

extern "C" __declspec(dllexport)
double xMultiply2Numbers(double x, double y)
{
	return x * y;
}

extern "C" __declspec(dllexport)
LPXLOPER12 xSolveTridag(LPXLOPER12 A_in, LPXLOPER12 b_in)
{
	FreeAllTempMemory();

	kMatrix<double> A;
	if(!kXlUtils::getMatrix(A_in, A))
		return TempStr12("input 1 is not a matrix");

	kVector<double> b;
	if(!kXlUtils::getVector(b_in, b))
		return TempStr12("input 2 is not a vector");

	if(b.size()!=A.rows())
		return TempStr12("input 1 and 2 must have number of rows");

	kVector<double> res, mem;
	kMatrixAlgebra::tridag(A, b, res, mem);

	LPXLOPER12 out = TempXLOPER12();
	kXlUtils::setVector(res, out);
	return out;
}

extern "C" __declspec(dllexport)
LPXLOPER12 xBanMul(LPXLOPER12 A_in, LPXLOPER12 b_in)
{
	FreeAllTempMemory();

	kMatrix<double> A;
	if(!kXlUtils::getMatrix(A_in, A))
		return TempStr12("input 1 is not a matrix");

	int m = A.cols()/2;
	if(A.cols()!=m*2+1)
		return TempStr12("input 1 must have odd number of columns");

	kVector<double> b;
	if(!kXlUtils::getVector(b_in, b))
		return TempStr12("input 2 is not a vector");

	if(b.size()!=A.rows())
		return TempStr12("input 1 and 2 must have number of rows");

	kVector<double> res;
	kMatrixAlgebra::banmul(A, m, m, b, res);

	LPXLOPER12 out = TempXLOPER12();
	kXlUtils::setVector(res, out);
	return out;
}

extern "C" __declspec(dllexport)
LPXLOPER12 xMatrixMul(LPXLOPER12 A_in, LPXLOPER12 B_in)
{
	FreeAllTempMemory();

	kMatrix<double> A;
	if(!kXlUtils::getMatrix(A_in, A))
		return TempStr12("input 1 is not a matrix");

	kMatrix<double> B;
	if(!kXlUtils::getMatrix(B_in, B))
		return TempStr12("input 2 is not a vector");

	if(B.rows()!=A.cols())
		return TempStr12("input 2 must have number of rows same as input 1 number of cols");

	kMatrix<double> res;
	kMatrixAlgebra::mmult(A, B, res);

	LPXLOPER12 out = TempXLOPER12();
	kXlUtils::setMatrix(res, out);
	return out;
}

//	Registers

extern "C" __declspec(dllexport) int xlAutoOpen(void)
{
	XLOPER12 xDLL;	

	Excel12f(xlGetName, &xDLL, 0);

	Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
		(LPXLOPER12)TempStr12(L"xMultiply2Numbers"),
		(LPXLOPER12)TempStr12(L"BBB"),
		(LPXLOPER12)TempStr12(L"xMultiply2Numbers"),
		(LPXLOPER12)TempStr12(L"x, y"),
		(LPXLOPER12)TempStr12(L"1"),
		(LPXLOPER12)TempStr12(L"myOwnCppFunctions"),
		(LPXLOPER12)TempStr12(L""),
		(LPXLOPER12)TempStr12(L""),
		(LPXLOPER12)TempStr12(L"Multiplies 2 numbers"),
		(LPXLOPER12)TempStr12(L""));
	
	Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
		(LPXLOPER12)TempStr12(L"xSolveTridag"),
		(LPXLOPER12)TempStr12(L"QQQ"),
		(LPXLOPER12)TempStr12(L"xSolveTridag"),
		(LPXLOPER12)TempStr12(L"A, b"),
		(LPXLOPER12)TempStr12(L"1"),
		(LPXLOPER12)TempStr12(L"myOwnCppFunctions"),
		(LPXLOPER12)TempStr12(L""),
		(LPXLOPER12)TempStr12(L""),
		(LPXLOPER12)TempStr12(L"Solving tri-diagonal system"),
		(LPXLOPER12)TempStr12(L""));
	
	Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
		(LPXLOPER12)TempStr12(L"xBanMul"),
		(LPXLOPER12)TempStr12(L"QQQ"),
		(LPXLOPER12)TempStr12(L"xBanMul"),
		(LPXLOPER12)TempStr12(L"A, b"),
		(LPXLOPER12)TempStr12(L"1"),
		(LPXLOPER12)TempStr12(L"myOwnCppFunctions"),
		(LPXLOPER12)TempStr12(L""),
		(LPXLOPER12)TempStr12(L""),
		(LPXLOPER12)TempStr12(L"Multiplying band matrix with a vector, assuming same number of lower and upper diags."),
		(LPXLOPER12)TempStr12(L""));
	
	Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
		(LPXLOPER12)TempStr12(L"xMatrixMul"),
		(LPXLOPER12)TempStr12(L"QQQ"),
		(LPXLOPER12)TempStr12(L"xMatrixMul"),
		(LPXLOPER12)TempStr12(L"A, B"),
		(LPXLOPER12)TempStr12(L"1"),
		(LPXLOPER12)TempStr12(L"myOwnCppFunctions"),
		(LPXLOPER12)TempStr12(L""),
		(LPXLOPER12)TempStr12(L""),
		(LPXLOPER12)TempStr12(L"Multiplying 2 matrices."),
		(LPXLOPER12)TempStr12(L""));
	
	/* Free the XLL filename */
	Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);

	return 1;
}

