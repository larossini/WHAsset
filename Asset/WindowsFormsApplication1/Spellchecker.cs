using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    class SpellChecker : Form1
    {
        string[] dictoinary = new string[] { "asset", "wo", "work order", "sn", "serial", "serial number", "po", "purchase order", "mac", "mac address" };

        public void spellchecker(string received)
        {
           
                // This will correct any minor spelling errors in the words before complairing them to the location
                Int32 textlength = (received.Length) - 1;//Grabs the length of received
                Int32 minlength = (textlength) - 1;//Sets a int for the least amount of character that can match
                string couldmatch = "";//Temp spot for the most likley match
                Int32 x = 0;//Int to walk through Array but not past Array length
                Int32 walkstart = 0;// Int to walk through received letter by letter
                Int32 ticker = 0;//Int ticker used to add the number of matched characters
                Int32 caught = 0;//Int caught used to hold the number of matched characters to compare to future passible matches
                bool match = false;//Bool to set if a match found
                Int32 closestwordspot = -1;//Int to hold the array index of closest 
                Int32 closestwordcount = 0;//Int used to hold number of matched characters
                received = received.ToLower();//Sends received to all lower characters
                while (((match != true) && (x <= 8)))
                //^^Wile no match found and X is less that Array indexes
                {
                    string dictemp = dictoinary[x];//Temp string set to array index
                    Int32 diclength = dictemp.Length;//Int set to length of DICTEMP string
                    if (((diclength > walkstart) && (walkstart <= textlength)))
                    //Checks if walk is less than the length of the dictionary string
                    {
                        if ((received.Substring(walkstart, 1) == dictemp.Substring(walkstart, 1)))
                        //Checks if INDIVIDUAL Characters from Dictionary and user input match
                        {
                            ticker = (ticker + 1);//If matched ticker adds 1
                            if ((walkstart <= diclength))//Checks if we can walk to the another character
                            {
                                walkstart = (walkstart + 1);//Walks to the righ a character
                            }

                        }
                        else if (((ticker >= minlength) && (ticker > caught)))
                        //^^If characters do not match and the ticker is longer than
                        //^^the last word
                        {
                            caught = ticker;//Caught becomes the new minimum to beat
                            couldmatch = dictemp;//Grabs this potential match string
                        }
                        else
                        {
                            if ((ticker >= closestwordcount))
                            //^^Checks the match characters to the old matched max
                            {
                                closestwordspot = x;//grabs the Index of the closest word
                            }

                            x = (x + 1);//Walks to the next word in the DIC array
                            ticker = 0;//Resets the match character ticker
                            walkstart = 0;//Resets this to start at index 0 at next word
                        }

                    }
                    else
                    {
                        if ((ticker == (textlength + 1)))//Checks if full word was matched
                        {
                            match = true;//Sets match to true
                            received = dictoinary[x];//Sets the temp to the found word
                        }

                        x = (x + 1);//Walks to the right in the array
                        ticker = 0;//Resets the ticker for match characters
                        walkstart = 0;//Resets the walk for the individual chartacters
                    }

                    if (((ticker >= minlength) && ((ticker > caught) && (ticker != (textlength + 1)))))

                    {
                        caught = ticker;
                        closestwordspot = x;
                    }

                }
                ManualEntry unknownentry = new ManualEntry();
                if (((match == true) || (couldmatch != "")))//^^Checks for a match
                {

                    received = dictoinary[x - 1].ToString();//sets temp to matched word
                    //unknownentry.manualentry(received);//Send matched word to a method that fills the label and string
                }
                else
                {
                    string possible = ("Did you mean " + dictoinary[closestwordspot]);
                    //^^Asks if the closest word is correct
                    string result = MessageBox.Show(possible, "Error", MessageBoxButtons.YesNo).ToString();
                    //^^Grabs user input from messagebox
                    if ((result == "yes"))//Checks if user selected Yes
                    {
                        received = dictoinary[x].ToString();//Sets the temp to the Array dictionary word
                        //unknownentry.manualentry(received);//Sends the word to the method that will set the labels and strings
                     
                }
                    else if ((result == "no"))//If answer
                    {

                    }
                    else
                    {
                        string tryagain = MessageBox.Show("Opps something went wrong lets try again" + "Please enter the location for this scan ", "Opps", MessageBoxButtons.OK).ToString();
                        //^^General failer messgae to try again becuase no match was found
                        spellchecker(tryagain);//Grabs new scan and self invokes
                    }

                }

            }//Used to check spelling of user input


        
    }
}
