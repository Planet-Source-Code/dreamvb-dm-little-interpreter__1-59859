ref:isOdd "Returns if 1 or 0 based on vlaue of ax";
ref:isZero "Returns 1 if value pf dx is zero";
ref:Log10 "Returns the base-10 logarithm of a number";

function isOdd(int ax)
{
   return ax % 2;
}

function isZero(int dx)
{
   return dx = 0;
}

function Log10(int cx)
{
   return log(cx) \ log(10)
}
