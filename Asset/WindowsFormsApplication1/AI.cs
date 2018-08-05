using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{
    class AI : Form1
    {
        int leng = -1;
        int bch = 0;
        

      
        public string Ai(string scan)//Scan contains characters
        {
            
            

            return "Not used yet";
        }

        public string Ai(long scan)//Scan contains all numerical values
        {
            leng = scan.ToString().Length;
            bch = currentBCHasset;
            if (leng == 6)//WO or PO
            {

                if (scan>=currentWO)//Its a WO
                {
                    return "wo";

                }
                else if(scan < currentWO)//Its a PO
                {
                    return "po";
                }
                else// Nothing matched
                {
                    return "help";
                }
             }else if (scan<bch && scan.ToString().Length==5)//WH Asset
            {
                 facility= 1;
                return "asset";
            }else if (scan>=bch && scan.ToString().Length == 5)//BCH Asset
            {
                facility = 3;
                return "asset";
                

            }
            else
            {
                ManualEntry unknownentry = new ManualEntry();
                string founditem = unknownentry.manualentry(scan);
                return founditem;
            }

        }
        public string Ai(Char scan)//Scan contains a uni character value
        {
            return "This is a char scan";

        }
    }
}
