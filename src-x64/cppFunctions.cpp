#include <Rcpp.h>
#include <iostream>
#include <fstream>
#include <sstream>

using namespace Rcpp;
using namespace std;

string itos(int i) // convert int to string
{
  stringstream s;
  s << i;
  return s.str();
}

// [[Rcpp::export]]
SEXP writeFile(std::string parent, std::string xmlText, std::string parentEnd, std::string R_fileName) {
  
  const char * s = R_fileName.c_str();
  
  std::ofstream xmlFile;
  xmlFile.open (s);
  xmlFile << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
  xmlFile << parent;
  xmlFile << xmlText;
  xmlFile << parentEnd;
  xmlFile.close();
  
  return 0;
}

