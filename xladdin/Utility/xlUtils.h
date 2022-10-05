#pragma once

#include "kMatrix.h"

namespace kXlUtils
{
    bool getVector(LPXLOPER12 in, kVector<double>& out);
    bool getMatrix(LPXLOPER12 in, kMatrix<double>& out);

    void setVector(const kVector<double>& in, LPXLOPER12 out);
    void setMatrix(const kMatrix<double>& in, LPXLOPER12 out);
}

bool kXlUtils::getVector(LPXLOPER12 in, kVector<double>& out)
{
    size_t m = getRows(in);

    out.resize(m);

    for(size_t i=0;i<m;++i)
    {
        double d = getNum(in, i, 0);
        if(d==numeric_limits<double>::infinity())
            return false;
        out((int)i) = d;
    }
    return true;
}

bool kXlUtils::getMatrix(LPXLOPER12 in, kMatrix<double>& out)
{
    size_t m = getRows(in);
    size_t n = getCols(in);
    out.resize(m, n);

    for(size_t i=0;i<m;++i)
    {
        for(size_t j=0;j<n;++j)
        {
            double d = getNum(in, i, j);
            if(d==numeric_limits<double>::infinity())
                return false;
            out((int)i, (int)j) = d;
        }
    }
    return true;
}

void kXlUtils::setVector(const kVector<double>& in, LPXLOPER12 out)
{
    size_t m = in.size();
    resize(out,m,1);
    for(size_t i=0;i<m;++i)
        setNum(out,in((int)i),i,0);
}

void kXlUtils::setMatrix(const kMatrix<double>& in, LPXLOPER12 out)
{
    size_t m = in.rows();
    size_t n = in.cols();
    resize(out,m,n);
    for(size_t i=0;i<m;++i)
        for(size_t j=0;j<n;++j)
            setNum(out,in((int)i,(int)j),i,j);
}
