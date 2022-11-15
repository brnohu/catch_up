#pragma once

//	desc:	1d finite difference solution for pdes of the form
//
//		0 = dV/dt + A V
//
//		A = -r + mu d/dx + 1/2 var d^2/dx^2
//
//	using the theta scheme
//
//		[1-theta dt A] V(t) = [1 + (1-theta) dt A] V(t+dt)
//

//	includes
#include "kMatrixAlgebra.h"

namespace kFiniteDifference
{
	//	1st order diff operator
	template <class V>
	void dx(
		const int			wind,
		const kVector<V>&	x,
		kMatrix<V>&			out)
	{
		int n = x.size()-1;
		if(n<1)
			return;

		V dxu   =  x(1)-x(0);
		out(0,0) =  0.0;
		out(0,1) = -1.0/dxu;
		out(0,2) =  1.0/dxu;

		for(int i = 1;i<n;++i)
		{
			V dxl = x(i)-x(i-1);
			V dxu = x(i+1)-x(i);
			if(wind<0)
			{
				out(i,0) = -1.0/dxl;
				out(i,1) =  1.0/dxl;
				out(i,2) =  0.0;
			}
			else if(wind==0)
			{
				out(i,0) = -dxu/dxl/(dxl+dxu);
				out(i,1) = (dxu/dxl-dxl/dxu)/(dxl+dxu);
				out(i,2) =  dxl/dxu/(dxl+dxu);
			}
			else
			{
				out(i,0) =  0.0;
				out(i,1) = -1.0/dxu;
				out(i,2) =  1.0/dxu;
			}
		}

		V dxl   =  x(n)-x(n-1);
		out(n,0) = -1.0/dxl;
		out(n,1) =  1.0/dxl;
		out(n,2) =  0.0;

		//	done
		return;
	}

	//	2nd order diff operator
	template <class V>
	void dxx(
		const kVector<V>&	x,
		kMatrix<V>&			out)
	{
		int n = x.size()-1;
		if(n<1)
			return;

		out(0,0) = V(0.0);
		out(0,1) = V(0.0);
		out(0,2) = V(0.0);

		for(int i = 1;i<n;++i)
		{
			V dxl	 = x(i)  -  x(i-1);
			V dxu	 = x(i+1) - x(i);
			out(i,0) =   2.0/(dxl*(dxl+dxu));
			out(i,1) = -(2.0/dxl+2.0/dxu)/(dxl+dxu);
			out(i,2) =   2.0/(dxu*(dxl+dxu));
		}

		out(n,0) = V(0.0);
		out(n,1) = V(0.0);
		out(n,2) = V(0.0);

		//	done
		return;
	}

}

//	class declaration
template <class V>
class kFd1d
{
public:

	//	init 
	void	init(
		int							numV,
		const kVector<V>&			x,
		bool						log);

	//	read
	const kVector<V>&				x()		const{ return myX; }
	const kVector<V>&				r()		const{ return myR; }
	const kVector<V>&				mu()	const{ return myMu; }
	const kVector<V>&				var()	const{ return myVar; }
	const kVector<kVector<V>>&		res()	const{ return myRes; }

	//	write
	kVector<V>&						x()		{ return myX; }
	kVector<V>&						r()		{ return myR; }
	kVector<V>&						mu()	{ return myMu; }
	kVector<V>&						var()	{ return myVar; }
	kVector<kVector<V>>&			res()	{ return myRes; }

	//	operator
	void	calcAx(
		const V&						one,
		const V&						dtTheta,
		int								wind,
		bool							tr,
		kMatrix<V>&						A) const;

	//	roll bwd
	void	rollBwd(
		const V&						dt,
		const V&						theta,
		int								wind,
		kVector<kVector<V>>&			res);

	//	roll fwd
	void	rollFwd(
		const V&						dt,
		const V&						theta,
		int								wind,
		kVector<kVector<V>>&			res);

	//	data
private:

	//	x, r, mu, var
	kVector<V>	myX, myR, myMu, myVar;

	//	log ?
	bool						myLog{false};

	//	helpers

	//	diff operators
	kMatrix<V>	myDxd,myDxu,myDx,myDxx;

	//	operator matrix
	kMatrix<V>					myA;
	kMatrix<V>					myAl;

	//	helper
	kVector<V>					myVs, myWs;

	//	vector of results
	kVector<kVector<V>>			myRes;
};

//	init
template <class V>
void	
kFd1d<V>::init(
	int							numV,
	const kVector<V>&			x,
	bool						log)
{
	myX = x;
	myRes.resize(numV);

	//	resize params
	myR.resize(myX.size());
	myMu.resize(myX.size());
	myVar.resize(myX.size());

	myR   = 0.0;
	myMu  = 0.0;
	myVar = 0.0;

	int numC = 3;

	myDxd.resize(myX.size(),numC);
	myDx .resize(myX.size(),numC);
	myDxu.resize(myX.size(),numC);
	myDxx.resize(myX.size(),numC);

	if(myX.empty()) return;

	kFiniteDifference::dx(-1,myX,myDxd);
	kFiniteDifference::dx( 0,myX,myDx);
	kFiniteDifference::dx( 1,myX,myDxu);
	kFiniteDifference::dxx(  myX,myDxx);

	//	log transform case
	myLog = log;
	if(myLog)
	{
		int n = myX.size()-1;
		for(int i = 1;i<n;++i)
		{
			for(int j=0;j<numC;++j)
				myDxx(i,j) -= myDx(i,j);
		}
	}

	myA.resize(myX.size(),numC);
	myAl.resize(myX.size(),numC);
	myVs.resize(myX.size());
	myWs.resize(myX.size());

	//	done
	return;
}

//	construct operator
template <class V>
void	
kFd1d<V>::calcAx(
	const V&		one,
	const V&		dtTheta,
	int				wind,
	bool			tr,
	kMatrix<V>&		A) const
{
	//	helps
	int i;
	V   aa;

	//	dims
	int n  = myX.size()-1;
	int mm = myDx.cols()/2;

	//	pad
	int pp = min(mm,myX.size());
	for(i=0;i<pp;++i)
	{
		for(int j=0;j<A.cols();++j)
		{
			A(i,j)   = V(0.0);
			A(n-i,j) = V(0.0);
		}
	}

	//	fill
	for(i=0;i<=n;++i)
	{
		int ml = max(0, i - mm);
		int mu = min(i + mm, n);
		int sign_mu = myMu(i)<0.0 ? -1 : (myMu(i)>0.0 ? 1 : 0);
		int dir = wind >= 2 ? sign_mu : wind;
		const kMatrix<V>& Dx = dir<0 ? myDxd : (!dir ? myDx : myDxu);
		for(int m = ml;m<=mu;++m)
		{
			int j = m - i + mm;
			aa = dtTheta*(myMu(i)*Dx(i,j) + 0.5*myVar(i)*myDxx(i,j));
			if(m==i) aa += one - dtTheta*myR(i);
			if(tr) A(m,i-m+mm) = aa;
			else   A(i,j)      = aa;
		}
	}

	//	done
	return;
}

//	roll bwd
template <class V>
void	
kFd1d<V>::rollBwd(
	const V&				dt,
	const V&				theta,
	int						wind,	
	kVector<kVector<V>>&	res)
{
	//	helps
	int k;
	
	//	dims
	int n    = myX.size();
	int mm   = myDx.cols()/2;
	int numV = (int)res.size();

	//	explicit
	if(theta!=1.0)
	{
		calcAx(1.0,dt*(1.0-theta),wind,false,myA);
		for(k=0;k<numV;++k)
		{
			myVs = res[k];
			kMatrixAlgebra::banmul(myA,mm,mm,myVs,res[k]);
		}
	}

	//	implicit
	if(theta!=0.0)
	{
		calcAx(1.0,-dt*theta,wind,false,myA);
		for(k=0;k<numV;++k)
		{
			myVs = res[k];
			kMatrixAlgebra::tridag(myA,myVs,res[k],myWs);
		}
	}

	//	done
	return;
} 

//	roll fwd
template <class V>
void
kFd1d<V>::rollFwd(
	const V&					dt,
	const V&					theta,
	int							wind,
	kVector<kVector<V>>&		res)
{
	//	helps
	int k;

	//	dims
	int n    = myX.size();
	int mm   = myDx.cols()/2;
	int numV = (int)res.size();

	//	implicit
	if(theta!=0.0)
	{
		calcAx(1.0,-dt*theta,wind,true,myA);
		for(k=0;k<numV;++k)
		{
			myVs = res[k];
			kMatrixAlgebra::tridag(myA,myVs,res[k],myWs);
		}
	}

	//	explicit
	if(theta!=1.0)
	{
		calcAx(1.0,dt*(1.0-theta),wind,true,myA);
		for(k=0;k<numV;++k)
		{
			myVs = res[k];
			kMatrixAlgebra::banmul(myA,mm,mm,myVs,res[k]);
		}
	}

	//	done
	return;
}

