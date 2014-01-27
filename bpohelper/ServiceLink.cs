using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace bpohelper
{
    struct ScreenSize
    {
        public int x;
        public int y;
    }

     struct Position
    {
         public int x;
         public int y;
    }
    class WebFormInputControl
    {
        public enum inputType { SelectionBox, TextBox };
        Dictionary<ScreenSize, Position> absolutePositionList;

        string defaultImageName;
        Position currentPostion;
        Position defaultPosition;
        
       

        public WebFormInputControl(inputType i, ScreenSize s, Position p)
        {

        }

        public WebFormInputControl()
        {

        }
        //public void InputData(inputType i, ScreenSize s, )
        //{

        //}
    }

    class ServiceLinkWebControl : WebFormInputControl
    {
        Position p;

        public void InputData()
        {

        }
    }

    class ServiceLinkInterface
    {
        public enum locationType { Urban, Suburban, Rural };
        public enum propertyType { SFR, Coop, Condo, PUD, MuliUnit, Manufactured, VacantLand, Other };
        public enum attDet { Attached, Detached };

        private ScreenSize screenSize;
        // private WebFormInputControl locationField;

        public ServiceLinkInterface()
        {
            // screenSize.x = 1600;
            // screenSize.y = 1000;
            // locationField = new WebFormInputControl(WebFormInputControl.inputType.SelectionBox, screenSize, 
            //pages.Add("subject", new ServiceLinkSubjectPage());

        }

        //private Dictionary<string, WebForm> pages;

        //StringBuilder macro = new StringBuilder();
        int timeout = 90;
        protected iMacros.App iim = new iMacros.App();

        string windowSize = @"SIZE X=1600 Y=1000";
       // string windowSize = @"SIZE X=947 Y=648";

        //Select box image
        // C:\Users\Scott\Documents\My Dropbox\Dev\bpohelper\bpohelper\20120811_2211.png

        //   macro.AppendLine(@"SIZE X=947 Y=648")
        public ServiceLinkInterface(iMacros.App d)
        {
            iim = d;
        }

        public void GetScreenPosition()
        {

        }

        public virtual void Location(locationType l)
        {

            StringBuilder macro = new StringBuilder();
            
            
            //find current location label* image and save its coordinates
            macro.AppendLine(windowSize);
            macro.AppendLine(@"IMAGESEARCH POS=1 IMAGE=20120811_2211.png CONFIDENCE=90");
            macro.AppendLine(@"ADD !EXTRACT {{!IMAGEX}}");
            macro.AppendLine(@"ADD !EXTRACT {{!IMAGEY}}");
            iim.iimPlayCode(macro.ToString(), timeout);
            string locationX = iim.iimGetExtract(1);
            string locationY = iim.iimGetExtract(2);

            string inputX = "";
            string inputY = "";
            macro.Clear();

            iMacros.Status status = iMacros.Status.sOk;
            //find correct Please Select box, ie same Y axis as image, and closest to field on x axis
            int i=1;
            Dictionary<int, Position> selectBoxList = new Dictionary<int, Position>();
            Position p = new Position();

            while (status == iMacros.Status.sOk)
            {
                macro.AppendLine(@"IMAGESEARCH POS=" + i.ToString() + " IMAGE=20120812_1728.png CONFIDENCE=90");
                macro.AppendLine(@"ADD !EXTRACT {{!IMAGEX}}");
                macro.AppendLine(@"ADD !EXTRACT {{!IMAGEY}}");
                status = iim.iimPlayCode(macro.ToString(), timeout);

                try
                {
                    p.x = Convert.ToInt16(iim.iimGetExtract(1));
                    p.y = Convert.ToInt16(iim.iimGetExtract(2));
                    selectBoxList.Add(i, p);
                }
                catch
                {
                }
               
                //inputY = iim.iimGetExtract(2);

                macro.Clear();
                i++;

                //if (Convert.ToInt16(locationY) - 15 <= Convert.ToInt16(inputY) && Convert.ToInt16(inputY) <= Convert.ToInt16(locationY) + 15)
                //{
                //    break;
                //}
            }
    
            //string inputX = iim.iimGetExtract(3);
            //string inputY = iim.iimGetExtract(4);
           
           

            iim.iimPlayCode(@"DS CMD=CLICK X=" + inputX + " Y=" + locationY + " CONTENT=" + l.ToString().ToLower()[0] + "{ENTER}");

            //switch ((int)l)
            //{
            //    case 0:
            //        macro.AppendLine(@"DS CMD=CLICK X={{!IMAGEX}} Y=330 CONTENT=u");
            //        break;
            //    case 1:
            //        iim.iimPlayCode(@"DS CMD=CLICK X=" + inputX + " Y=" + locationY + " CONTENT=s{ENTER}");
            //        break;
            //    case 2:
            //        macro.AppendLine(@"DS CMD=CLICK X={{!IMAGEX}} Y=330 CONTENT=r");
            //        break;

            //}

            // macro.AppendLine(@"WAIT SECONDS=10");
            // macro.AppendLine(@"IMAGESEARCH POS=1 IMAGE=20120812_1728.png CONFIDENCE=95");
            // macro.AppendLine(@"IMAGECLICK POS=1 IMAGE=20120812_1728.png CONFIDENCE=95 ");
            // macro.AppendLine(@"WAIT SECONDS=10");
            // macro.AppendLine(@"ADD !EXTRACT {{!IMAGEX}}");
            // macro.AppendLine(@"PROMPT  {{!IMAGEX}}");
            // macro.AppendLine(@"WAIT SECONDS=3");
            // macro.AppendLine(@"DS CMD=KEY X=0 Y=0 CONTENT={ENTER}");
            // macro.AppendLine(@"WAIT SECONDS=5");
            //// macro.AppendLine(@"DS CMD=CLICK X=460 Y=374 CONTENT=");
            //// macro.AppendLine(@"WAIT SECONDS=5");
            // macro.AppendLine(@"DS CMD=CLICK X=460 Y=419 CONTENT=");
            // macro.AppendLine(@"WAIT SECONDS=3");
            // macro.AppendLine(@"DS CMD=KEY X=0 Y=0 CONTENT={ENTER}");
            // macro.AppendLine(@"WAIT SECONDS=1");
            // macro.AppendLine(@"DS CMD=CLICK X=460 Y=449 CONTENT=");
            // macro.AppendLine(@"WAIT SECONDS=3");
            // macro.AppendLine(@"DS CMD=KEY X=0 Y=0 CONTENT={ENTER}");
            // macro.AppendLine(@"WAIT SECONDS=1");
            // macro.AppendLine(@"DS CMD=CLICK X=460 Y=494 CONTENT=");
            // macro.AppendLine(@"WAIT SECONDS=3");
            // macro.AppendLine(@"DS CMD=KEY X=0 Y=0 CONTENT={ENTER}");



            //int x = 0;

        }

        public void PropertyType(propertyType p)
        {
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(windowSize);
            //macro.AppendLine(@"DS CMD=CLICK X=458 Y=543 CONTENT=");
            //macro.AppendLine(@"WAIT SECONDS=1");
            // macro.AppendLine(@"DS CMD=KEY CONTENT=sccmmmmamm");
            switch ((int)p)
            {
                case 0:
                    macro.AppendLine(@"DS CMD=CLICK X=458 Y=543 CONTENT=s");
                    break;
                case 1:
                    macro.AppendLine(@"DS CMD=CLICK X=458 Y=543 CONTENT=c");
                    break;
                case 2:
                    macro.AppendLine(@"DS CMD=CLICK X=458 Y=543 CONTENT=cc");
                    break;

            }
            macro.AppendLine(@"DS CMD=KEY X=0 Y=0 CONTENT={ENTER}");
            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, timeout);
        }

        public void FillField(string fieldName, object value)
        {
             StringBuilder macro = new StringBuilder();
             iMacros.Status status = iMacros.Status.sOk;
             string locationX = "";
             string locationY = "";
            
            //find current location label* image and save its coordinates
            macro.AppendLine(windowSize);
            macro.AppendLine(@"IMAGESEARCH POS=1 IMAGE=" + fieldName + " CONFIDENCE=90");
            macro.AppendLine(@"ADD !EXTRACT {{!IMAGEX}}");
            macro.AppendLine(@"ADD !EXTRACT {{!IMAGEY}}");
            status = iim.iimPlayCode(macro.ToString(), timeout);
           

            if (status == iMacros.Status.sOk)
            {
                locationX = iim.iimGetExtract(1);
                locationY = iim.iimGetExtract(2);

                string path = @"C:\Users\Scott\Documents\My Dropbox\Dev\bpohelper\ServiceLink\imageLocationInfo.txt";

                if (!File.Exists(path))
                {
                    // Create a file to write to.
                    using (StreamWriter sw = File.CreateText(path))
                    {
                        sw.WriteLine("{0}-->X={1}, Y={2}", fieldName, locationX, locationY);
                    }
                }

                // This text is always added, making the file longer over time
                // if it is not deleted.
                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine("{0}-->X={1}, Y={2}", fieldName, locationX, locationY);
                }	

            }

            string inputX = "";
            string inputY = "";
            macro.Clear();

             status = iMacros.Status.sOk;
            //find correct Please Select box, ie same Y axis as image, and closest to field on x axis
            int i=1;
            Dictionary<int, Position> selectBoxList = new Dictionary<int, Position>();
            Position p = new Position();

            int xDiff = 5000;
            inputX = "0";
            while (status == iMacros.Status.sOk)
            {
                macro.AppendLine(@"IMAGESEARCH POS=" + i.ToString() + " IMAGE=selectionBox-PleaseSelect.png CONFIDENCE=90");
                macro.AppendLine(@"ADD !EXTRACT {{!IMAGEX}}");
                macro.AppendLine(@"ADD !EXTRACT {{!IMAGEY}}");
                status = iim.iimPlayCode(macro.ToString(), timeout);

                try
                {
                    p.x = Convert.ToInt16(iim.iimGetExtract(1));
                    p.y = Convert.ToInt16(iim.iimGetExtract(2));
                    selectBoxList.Add(i, p);
                }
                catch
                {
                }
               
                //inputY = iim.iimGetExtract(2);

                macro.Clear();
                i++;

                if (Convert.ToInt16(locationY) - 15 <= p.y && Convert.ToInt16(locationY) <= p.y + 15)
                {
                     inputY = iim.iimGetExtract(2);
                }

                if (Math.Abs((Convert.ToInt16(locationX)  - p.x )) < xDiff)
                {
                    inputX = p.x.ToString();
                    xDiff = Math.Abs((Convert.ToInt16(locationX) - p.x));
                }
                
            }

    
            //string inputX = iim.iimGetExtract(3);
            //string inputY = iim.iimGetExtract(4);
           
           

            iim.iimPlayCode(@"DS CMD=CLICK X=" + inputX + " Y=" + locationY + " CONTENT=" + value.ToString().ToLower()[0] + "{ENTER}");

        }
    }


    //class WebForm
    //{
    //    // protected ScreenSize defaultSize;
    //    protected iMacros.App iim;

    //    public WebForm(iMacros.App d)
    //    {
    //        iim = d;
    //    }

    //}

    //class ServiceLinkWebForm : WebForm
    //{

    //    //header does not resize
    //    //at this screen size, macro.AppendLine(@"SIZE X=947 Y=648"); 
    //    //vendor columme lines up with second data column to get X axis cood
    //    //vendor position
    //    //macro.AppendLine(@"DS CMD=CLICK X=489 Y=148 CONTENT=");

    //    //ServiceLinkWebControl location;

    //    // public void 


    //}

    class ServiceLinkSubjectPage : ServiceLinkInterface
    {
        Dictionary<string, string> detachedFieldList;
        Dictionary<string, string> attachedFieldList;
        // public ServiceLinkSubjectPage(iMacros.App d)
        //{
        //    iim = d;
        //    this.ServiceLinkSubjectPage();
        //}
        public ServiceLinkSubjectPage(iMacros.App d)
        {
            iim = d;
             Dictionary<string, string> commonFields = new Dictionary<string,string>() {
                 {"location", "20120811_2211.png"}, 
                 {"propType", "propType-star.png"}, 
                 {"attDet", "attDet-star.png"},
                 {"occStatus", "attDet-star.png"}


             };

            attachedFieldList = new Dictionary<string, string>();
            attachedFieldList = commonFields;
            attachedFieldList.Add("hoaFees", "hoa-fees.png");
            attachedFieldList.Add("totalUnitsInDev", "development-units.png");

        }

        public void FillPage(Dictionary<string, string> fieldValueList)
        {
            foreach (string k in fieldValueList.Keys)
            {
                //find the field on the screen
                //find the input control for the field
                //set focus to that control
                FillField(attachedFieldList[k], fieldValueList[k]);

                //fill the field with the value
                


            }

        }



    
    }
    class Screen
    {

    }

}
