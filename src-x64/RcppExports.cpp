// This file was generated by Rcpp::compileAttributes
// Generator token: 10BE3573-1514-4C36-9D1C-5A225CD40393

#include <Rcpp.h>

using namespace Rcpp;

// writeFile
SEXP writeFile(std::string parent, std::string xmlText, std::string parentEnd, std::string R_fileName);
RcppExport SEXP opendocx_writeFile(SEXP parentSEXP, SEXP xmlTextSEXP, SEXP parentEndSEXP, SEXP R_fileNameSEXP) {
BEGIN_RCPP
    SEXP __sexp_result;
    {
        Rcpp::RNGScope __rngScope;
        Rcpp::traits::input_parameter< std::string >::type parent(parentSEXP );
        Rcpp::traits::input_parameter< std::string >::type xmlText(xmlTextSEXP );
        Rcpp::traits::input_parameter< std::string >::type parentEnd(parentEndSEXP );
        Rcpp::traits::input_parameter< std::string >::type R_fileName(R_fileNameSEXP );
        SEXP __result = writeFile(parent, xmlText, parentEnd, R_fileName);
        PROTECT(__sexp_result = Rcpp::wrap(__result));
    }
    UNPROTECT(1);
    return __sexp_result;
END_RCPP
}
