using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{
    public class ManualEntry : Form1

    {
        string tempstring = "null";
        
        public string manualentry(long received)
        {

          string useranswer =  InputBox.Show("The scanned item can not be detirmed. Please tell me what information was entered.").Text;
            return "null";

        }
    }
}
