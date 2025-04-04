using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BMI.AML
{
    class cl_aletler
    {
       public string baglan = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";
       public string tarixgiris = "";
       public string tarixcixis = "";
       public DataTable dtsorgu = new DataTable();
    }
}
