using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{
    class Grab
    {
        public string localstring { get; }
        public long locallong { get; }
        public int localint { get; }
        public string location { get;}
        public string answer { get;  }
        public int facilitycode { get;  }
        int x;
        long xx;
        AI scan = new AI();

        public void grab()
        {
            
        }
        

        public Grab(string scannedtemp)
        {
            answer = "null";
            this.localstring = scannedtemp;
            bool testing = long.TryParse(localstring, out xx);// Tests to see if this scan was all numbers
            if (testing == true)
            {
                try{
                    int x1 = x;
                   x= checked((int)xx);
                }
                catch
                {
                    locallong = xx;
                    answer = scan.Ai(xx);
                    location = answer;
                }

                localint = x;
               answer= scan.Ai(x);// Call AI with INT
                location = answer;
            }
            else
            {
                
               answer = scan.Ai(localstring);// Call AI with String 
                location = answer;

            }



        }
    }
}
