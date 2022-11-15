
#include <windows.h>
#include "xlcall.h"
#include "framework.h"
#include "xlOper.h"
#include "../Utility/kMatrixAlgebra.h"
#include "../Utility/kFd1d.h"
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

extern "C" __declspec(dllexport)
LPXLOPER12 xFd1d(double t, LPXLOPER12 x_in, LPXLOPER12 r_in, LPXLOPER12 mu_in, LPXLOPER12 sigma_in, LPXLOPER12 v0_in, LPXLOPER12 tech_in)
{
	FreeAllTempMemory();

	kVector<double> tech;
	if(!kXlUtils::getVector(tech_in, tech))
		return TempStr12("input 1 is not a matrix");

	//	standard data
	int    numt  = tech.size()>0 ? (int)(tech(0)+0.5) : 1;
	double theta = tech.size()>1 ? tech(1) : 0.5;
	int    fb    = tech.size()>2 ? (int)(tech(2)+0.5) : -1;
	int    log   = tech.size()>3 ? (int)(tech(3)+0.5) :  0;
	int    wind  = tech.size()>4 ? (int)(tech(4)+0.5) :  0;

	//	fd grid
	kFd1d<double> fd;
	if(!kXlUtils::getVector(x_in, fd.x()))
		return TempStr12("input 1 is not a matrix");

	//	init fd
	fd.init(1,fd.x(),log>0);

	if(!kXlUtils::getVector(r_in, fd.r()))
		return TempStr12("input 1 is not a matrix");

	if(!kXlUtils::getVector(mu_in, fd.mu()))
		return TempStr12("input 2 is not a vector");

	if(!kXlUtils::getVector(sigma_in, fd.var()))
		return TempStr12("input 2 is not a vector");

	auto& fd_var = fd.var();
	for(int i=0;i<fd_var.size();++i)
		fd_var(i) *= fd_var(i);

	if(!kXlUtils::getVector(v0_in, fd.res()[0]))
		return TempStr12("input 2 is not a vector");

	if(theta<0.0) theta=0.0;
	if(theta>1.0) theta=1.0;

	double dt = t/numt;
	for(int n=0;n<numt;++n)
	{
		if(fb<=0)
		{
			fd.rollBwd(dt,theta,wind,fd.res());
		}
		else
		{
			fd.rollFwd(dt,theta,wind,fd.res());
		}
	}

	LPXLOPER12 out = TempXLOPER12();
	kXlUtils::setVector(fd.res()[0], out);
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

	Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
		(LPXLOPER12)TempStr12(L"xFd1d"),
		(LPXLOPER12)TempStr12(L"QBQQQQQQ"),
		(LPXLOPER12)TempStr12(L"xFd1d"),
		(LPXLOPER12)TempStr12(L"t, x, r, mu, sigma, v0, tech"),
		(LPXLOPER12)TempStr12(L"1"),
		(LPXLOPER12)TempStr12(L"myOwnCppFunctions"),
		(LPXLOPER12)TempStr12(L""),
		(LPXLOPER12)TempStr12(L""),
		(LPXLOPER12)TempStr12(L"Solve 1d fd."),
		(LPXLOPER12)TempStr12(L""));

	/* Free the XLL filename */
	Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);

	return 1;
}

