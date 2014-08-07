using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace bpohelper
{

    //
    //based on form REOamerica
    //

    class Dispo
    {

        private string formType = "Standard";

        public void UploadPics(Form1 form, iMacros.App iim)
        {
            string fullfFilePath = form.SubjectFilePath.Replace(" ", "<SP>");


            StringBuilder macro = new StringBuilder();
            //comps
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:valbpophotouploaddriveby.aspx ATTR=VALUE:UPLOAD<SP>PHOTO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:PhotoUpload ATTR=ID:fileImage CONTENT=" + fullfFilePath + "\\L1.jpg");


            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:PhotoUpload ATTR=ID:btnUpload&&VALUE:UPLOAD<SP>THIS<SP>PHOTO");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:SUBMIT FORM=ACTION:valbpophotouploaddriveby.aspx ATTR=VALUE:UPLOAD<SP>PHOTO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:PhotoUpload ATTR=ID:fileImage CONTENT=" + fullfFilePath + "\\L2.jpg");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:PhotoUpload ATTR=ID:btnUpload&&VALUE:UPLOAD<SP>THIS<SP>PHOTO");
            macro.AppendLine(@"TAG POS=3 TYPE=INPUT:SUBMIT FORM=ACTION:valbpophotouploaddriveby.aspx ATTR=VALUE:UPLOAD<SP>PHOTO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:PhotoUpload ATTR=ID:fileImage CONTENT=" + fullfFilePath + "\\L3.jpg");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:PhotoUpload ATTR=ID:btnUpload&&VALUE:UPLOAD<SP>THIS<SP>PHOTO");
            macro.AppendLine(@"TAG POS=4 TYPE=INPUT:SUBMIT FORM=ACTION:valbpophotouploaddriveby.aspx ATTR=VALUE:UPLOAD<SP>PHOTO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:PhotoUpload ATTR=ID:fileImage CONTENT=" + fullfFilePath + "\\S1.jpg");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:PhotoUpload ATTR=ID:btnUpload&&VALUE:UPLOAD<SP>THIS<SP>PHOTO");
            macro.AppendLine(@"TAG POS=5 TYPE=INPUT:SUBMIT FORM=ACTION:valbpophotouploaddriveby.aspx ATTR=VALUE:UPLOAD<SP>PHOTO");
            macro.AppendLine(@"TAG POS=1 TYPE=TABLE FORM=ID:PhotoUpload ATTR=ID:Table1");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:PhotoUpload ATTR=ID:fileImage CONTENT=" + fullfFilePath + "\\S2.jpg");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:PhotoUpload ATTR=ID:btnUpload&&VALUE:UPLOAD<SP>THIS<SP>PHOTO");
            macro.AppendLine(@"TAG POS=6 TYPE=INPUT:SUBMIT FORM=ACTION:valbpophotouploaddriveby.aspx ATTR=VALUE:UPLOAD<SP>PHOTO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:PhotoUpload ATTR=ID:fileImage CONTENT=" + fullfFilePath + "\\S3.jpg");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:PhotoUpload ATTR=ID:btnUpload&&VALUE:UPLOAD<SP>THIS<SP>PHOTO");
            //subject
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:valbpophotouploaddriveby.aspx ATTR=VALUE:UPLOAD<SP>MAIN<SP>PHOTO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:PhotoUpload ATTR=ID:fileImage CONTENT=" + fullfFilePath + "\\front.jpg");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:PhotoUpload ATTR=ID:btnUpload&&VALUE:UPLOAD<SP>THIS<SP>PHOTO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ACTION:valbpoconditionphotouploaddriveby.aspx ATTR=VALUE:UPLOAD<SP>PHOTO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:PhotoUpload ATTR=ID:fileImage CONTENT=" + fullfFilePath + "\\addess.jpg");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:PhotoUpload ATTR=ID:btnUpload&&VALUE:UPLOAD<SP>THIS<SP>PHOTO");
            macro.AppendLine(@"TAG POS=2 TYPE=INPUT:SUBMIT FORM=ACTION:valbpoconditionphotouploaddriveby.aspx ATTR=VALUE:UPLOAD<SP>PHOTO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:FILE FORM=ID:PhotoUpload ATTR=ID:fileImage CONTENT=" + fullfFilePath + "\\street.jpg");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:PhotoUpload ATTR=ID:btnUpload&&VALUE:UPLOAD<SP>THIS<SP>PHOTO");
            string macroCode = macro.ToString();
           iim.iimPlayCode(macroCode, 60);

        }


        public string ReportType(iMacros.App iim)
        {
            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"FRAME NAME=main");
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=NAME:BPOForm ATTR=NAME:SpecialInst EXTRACT=TXT");
            iim.iimPlayCode(macro.ToString(), 30);
            string reportType = iim.iimGetLastExtract();
            if (reportType.Contains("Fannie Mae"))
            {
                formType = "z";
                return "z";
            }
            else
            {
                formType = "x";
                return "x";
            }
        }

        public void Prefill(iMacros.App iim, Form1 form)
        {
            StringBuilder macro = new StringBuilder();
            Dictionary<string, string> fieldList = new Dictionary<string, string>();
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");

            //
            //CheckBoxes
            //
            //Data Source
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:SubjectProperty_InformationSources_UsedTaxRecords&&VALUE:true CONTENT=YES");
            //Roof
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_RoofingMaterials_HasAsphalt&&VALUE:true CONTENT=YES");
            //Siding
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_ExteriorBuildingMaterials_HasAluminumOrVinyl&&VALUE:true CONTENT=YES");
            //BasementType
            if (form.SubjectBasementType.ToLower().Contains("full"))
            {
                 macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_Basement_IsFull&&VALUE:true CONTENT=YES");
            }
            if (form.SubjectBasementType.ToLower().Contains("partial"))
            {
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_Basement_IsPartial&&VALUE:true CONTENT=YES");
            }
            if (form.SubjectBasementType.ToLower().Contains("walkout"))
            {
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_Basement_IsWalkout&&VALUE:true CONTENT=YES");
            }
            if (form.SubjectBasementType.ToLower().Contains("slab"))
            {
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_Basement_IsSlab&&VALUE:true CONTENT=YES");
            }
            if (form.SubjectBasementType.ToLower().Contains("crawl"))
            {
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_Basement_HasCrawlspace&&VALUE:true CONTENT=YES");
            }
            if (!string.IsNullOrEmpty(form.SubjectNumberFireplaces) | form.SubjectNumberFireplaces != "0")
            {
                macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_Amenities_HasFireplace&&VALUE:true CONTENT=YES");
            }

            //
            //RadioButtons
            //
            //Currently Listed
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:SubjectProperty_LastListing_IsCurrentlyListed&&VALUE:true CONTENT=NO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:SubjectProperty_LastListing_IsCurrentlyListed&&VALUE:false CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:SubjectProperty_LastListing_WasListedInPastYear&&VALUE:true CONTENT=NO");
            //list past 12 months
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:SubjectProperty_LastListing_WasListedInPastYear&&VALUE:false CONTENT=YES");
            //Overall Rating
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:SubjectProperty_Ratings_GeneralConditionRatingType&&VALUE:Average CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:SubjectProperty_Ratings_ConstructionQualityType&&VALUE:Average CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:SubjectProperty_Ratings_ConformityRatingType&&VALUE:2 CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:SubjectProperty_Ratings_CurbAppealRatingType&&VALUE:2 CONTENT=YES");
            //Inventory
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:SubjectProperty_AreaMarketConditions_PropertyInventoryType&&VALUE:1 CONTENT=NO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:SubjectProperty_AreaMarketConditions_PropertyInventoryType&&VALUE:2 CONTENT=YES");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:SubjectProperty_AreaMarketConditions_PropertyInventoryType&&VALUE:1 CONTENT=NO");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:SubjectProperty_AreaMarketConditions_PropertyInventoryType&&VALUE:2 CONTENT=YES");
 
            
            //
            //Selection Boxes
            //
            //Occupancy
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:SubjectProperty_Occupancy_OccupancyStatusType CONTENT=%Unknown");
            //LocationTYpe
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:SubjectProperty_Site_LocationType CONTENT=%Suburban");
            //Primary Location Factor
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:SubjectProperty_Site_PrimaryLocationFactorType CONTENT=%1");
            //Neighborhood like property value trend
            macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:SubjectProperty_NeighborhoodMarketConditions_ValueTrendType CONTENT=%Improving");

            if (form.SubjectMlsType.ToLower().Contains("1 story"))
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_StyleType CONTENT=%1<SP>Story");
            }
            else if (form.SubjectMlsType.ToLower().Contains("2 stories"))
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_StyleType CONTENT=%2<SP>Story");
            }
            else if (form.SubjectMlsType.ToLower().Contains("raised ranch"))
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_StyleType CONTENT=%Muiltilevel");
            }
            else if (form.SubjectMlsType.ToLower().Contains("split") || form.SubjectMlsType.ToLower().Contains("other"))
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_StyleType CONTENT=%Muiltilevel");
            }
            else if (form.SubjectMlsType.ToLower().Contains("1.5"))
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_StyleType CONTENT=%1.5<SP>Story");
            }

            if (form.SubjectParkingType.ToLower().Contains("att"))
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_Garages_AttachedCarCountType CONTENT=%" + Regex.Match(form.SubjectParkingType, @"\d").Value + "<SP>cars");
            }
            else if (form.SubjectParkingType.ToLower().Contains("det"))
            {
                macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_Garages_DetachedCarCountType CONTENT=%" + Regex.Match(form.SubjectParkingType, @"\d").Value + "<SP>cars");
            }

            
           
           


            //
            //Textboxes
            //
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:Property_Address_County CONTENT=" + form.SubjectCounty);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_SubdivisionName CONTENT=" + form.SubjectSubdivision);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_ParcelNumber CONTENT=" + form.SubjectPin);


            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SuggestedPricing_AsIsQuickSaleListingPrice CONTENT=" + form.SubjectMarketValue);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SuggestedPricing_RepairedQuickListingPrice CONTENT=" + form.SubjectMarketValue);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SuggestedPricing_AsIsQuickSaleSellingPrice CONTENT=" + form.SubjectMarketValue);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SuggestedPricing_RepairedQuickSellingPrice CONTENT=" + form.SubjectMarketValue);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_HomeOwnerAssociation_FeeAmount CONTENT=0");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SuggestedPricing_AsIsListingPrice CONTENT=" + form.SubjectMarketValue);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SuggestedPricing_RepairedListingPrice CONTENT=" + form.SubjectMarketValue);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SuggestedPricing_AsIsSellingPrice CONTENT=" + form.SubjectMarketValue);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SuggestedPricing_RepairedSellingPrice CONTENT=" + form.SubjectMarketValue);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_GradedRooms_TotalRoomCount CONTENT=" + form.SubjectRoomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_GradedRooms_BedroomCount CONTENT=" + form.SubjectBedroomCount);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_GradedRooms_FullBathCount CONTENT=" + form.SubjectBathroomCount[0] );
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_GradedRooms_PowderRoomCount CONTENT=" + form.SubjectBathroomCount[2]);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_YearBuilt CONTENT=" + form.SubjectYearBuilt);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_Site_LotSize CONTENT=" + form.SubjectLotSize);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_Basement_SquareFeet CONTENT=" + form.SubjectBasementGLA);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_Basement_PercentFinished CONTENT=" + form.SubjectBasementFinishedGLA);
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_AreaMarketConditions_NormalMarketingDays CONTENT=90");
            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_LivingAreaSquareFeet CONTENT=" + form.SubjectAboveGLA);

            //
            //TextAreas
            //
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:form0 ATTR=ID:SubjectProperty_Comments CONTENT=Searched<SP>a<SP>distance<SP>of<SP>at<SP>least<SP>1<SP>miles,<SP>up<SP>to<SP>6<SP>months<SP>in<SP>time.<SP>The<SP>comps<SP>bracket<SP>the<SP>subject<SP>in<SP>age,<SP>SF,<SP>and<SP>lot<SP>size,<SP>as<SP>well<SP>as<SP>use<SP>comps<SP>in<SP>same<SP>condition<SP>with<SP>similar<SP>features,<SP>and<SP>from<SP>the<SP>subjects<SP>market<SP>area.<SP>All<SP>the<SP>comps<SP>are<SP>Reasonable<SP>substitute<SP>for<SP>the<SP>subject<SP>property,<SP>similar<SP>in<SP>most<SP>areas.<SP>Price<SP>opinion<SP>was<SP>based<SP>off<SP>comparable<SP>statistics.");
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:form0 ATTR=ID:SubjectProperty_NeighborhoodMarketConditions_IssuesDescription CONTENT=No<SP>known<SP>issues<SP>noted<SP>during<SP>drive-by<SP>inspection.");
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:form0 ATTR=ID:SubjectProperty_AreaMarketConditions_Comments CONTENT=No<SP>known<SP>special<SP>concerns,<SP>encroachments,<SP>easements,<SP>water<SP>rights,<SP>environmental<SP>concerns,<SP>flood<SP>zones.");
            macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:form0 ATTR=ID:SubjectProperty_AdditionalComments CONTENT=Searched<SP>a<SP>distance<SP>of<SP>at<SP>least<SP>1<SP>miles,<SP>up<SP>to<SP>6<SP>months<SP>in<SP>time.<SP>The<SP>comps<SP>bracket<SP>the<SP>subject<SP>in<SP>age,<SP>SF,<SP>and<SP>lot<SP>size,<SP>as<SP>well<SP>as<SP>use<SP>comps<SP>in<SP>same<SP>condition<SP>with<SP>similar<SP>features,<SP>and<SP>from<SP>the<SP>subjects<SP>market<SP>area.<SP>All<SP>the<SP>comps<SP>are<SP>Reasonable<SP>substitute<SP>for<SP>the<SP>subject<SP>property,<SP>similar<SP>in<SP>most<SP>areas.<SP>Price<SP>opinion<SP>was<SP>based<SP>off<SP>comparable<SP>statistics.");



         
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_Amenities_HasDeckPatio&&VALUE:true CONTENT=YES");
           // macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_Amenities_HasGuestHouse&&VALUE:true CONTENT=YES");
            //macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:SubjectProperty_Structure_Amenities_PoolType CONTENT=$Select...");
           
         
            
          
            
            
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_NeighborhoodMarketConditions_ValueTrendAnnualPercent CONTENT=6");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_AreaMarketConditions_HighSalePrice CONTENT=1000");
            //macro.AppendLine(@"TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:SubjectProperty_AreaMarketConditions_LowSalePrice CONTENT=1");
        
         
            //some chages on form 8/6/14



            string macroCode = macro.ToString();
            iim.iimPlayCode(macroCode, 50);


        }

        public void CompFill(iMacros.App iim, string saleOrList, string compNum, Dictionary<string, string>fieldList)
        {
            string[] compList = new string[6] { "0", "1", "2", "3", "4", "5" };
            string sol;
            if (saleOrList == "sale")
            {
                sol = "Sales";
            }
            else 
            {
                sol = "Listing";
            }
            StringBuilder move = new StringBuilder();
            move.AppendLine(@"SET !ERRORIGNORE YES");
            move.AppendLine(@"SET !REPLAYSPEED FAST");
            move.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:SUBJECT<SP>PROPERTY");

            if (sol == "Sales")
            {
                switch (compNum)
                {
                    case "0":
                        move.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:SALE<SP>COMPS");
                        break;
                    case "1":
                        move.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:SALE<SP>COMPS");
                        move.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=ID:ComparableOrRentalSelect CONTENT=%4");
                        break;
                    case "2":
                        move.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:SALE<SP>COMPS");
                        move.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=ID:ComparableOrRentalSelect CONTENT=%5");
                        break;
                }
            }
            else
            {
                switch (compNum)
                {
                    case "0":
                        move.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:LISTING<SP>COMPS");
                        break;
                    case "1":
                        move.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:LISTING<SP>COMPS");
                        move.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=ID:ComparableOrRentalSelect CONTENT=%1");
                        break;
                    case "2":
                        move.AppendLine(@"TAG POS=1 TYPE=A ATTR=TXT:LISTING<SP>COMPS");
                        move.AppendLine(@"TAG POS=1 TYPE=SELECT ATTR=ID:ComparableOrRentalSelect CONTENT=%2");
                        break;

                }

            }
          

            string moveCode = move.ToString();
            iim.iimPlayCode(moveCode, 60);


            StringBuilder macro = new StringBuilder();
            macro.AppendLine(@"SET !ERRORIGNORE YES");
            macro.AppendLine(@"SET !TIMEOUT_STEP 0");
            macro.AppendLine(@"SET !REPLAYSPEED FAST");

           foreach (string field in fieldList.Keys)
           {
               if (field.Contains("*"))
               {
                   //drop down box
                   foreach (string s in compList)
                   {
//                       macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:Comparables_3__Structure_Garages_DetachedCarCountType CONTENT=%2<SP>cars");

                     // macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:Comparables_2__Structure_StyleType CONTENT=%Multilevel");
                       //ClosingDetails_ClosingDate
                       macro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:Comparables_{0}__{1} CONTENT=%{2}\r\n", s, field.Replace("*", ""), fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                      // macro.AppendFormat("TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:Comparables_{0}__{1} CONTENT=%{2}\r\n", s, field.Replace("*", ""), fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                   }
                   
                  
                   
                   
               }
               else
               {
                    //0
                   macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:{2}Comparables_{0}__{1} CONTENT={3}\r\n", "0", field, sol, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                   macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:Comparables_{0}__{1} CONTENT={2}\r\n", "0", field, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                   //1
                   macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:{2}Comparables_{0}__{1} CONTENT={3}\r\n", "1", field, sol, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                   macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:Comparables_{0}__{1} CONTENT={2}\r\n", "1", field, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                  //2
                   macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:{2}Comparables_{0}__{1} CONTENT={3}\r\n", "2", field, sol, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                   macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:Comparables_{0}__{1} CONTENT={2}\r\n", "2", field, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
               //3
                 macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:{2}Comparables_{0}__{1} CONTENT={3}\r\n", "3", field, sol, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                   macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:Comparables_{0}__{1} CONTENT={2}\r\n", "3", field, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));

               //4
                 macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:{2}Comparables_{0}__{1} CONTENT={3}\r\n", "4", field, sol, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                   macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:Comparables_{0}__{1} CONTENT={2}\r\n", "4", field, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));

               //5
                 macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:{2}Comparables_{0}__{1} CONTENT={3}\r\n", "5", field, sol, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));
                   macro.AppendFormat("TAG POS=1 TYPE=INPUT:TEXT FORM=ID:form0 ATTR=ID:Comparables_{0}__{1} CONTENT={2}\r\n", "5", field, fieldList[field].Replace(",", "").Replace("$", "").Replace(" ", "<SP>"));

               }
           }

            //
            //Checkboxes
            //
            //Siding



       

           foreach (string s in compList)
           {
               compNum = s;

               macro.AppendFormat("TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:Comparables_{0}__Structure_ExteriorMaterials_HasAluminumOrVinyl&&VALUE:true CONTENT={1}\r\n", compNum, "YES");
               //BasementType
               if (fieldList["BasementType"].ToLower().Contains("full"))
               {
                   macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:Comparables_" + compNum + "__Structure_Basement_IsFull&&VALUE:true CONTENT=YES");
               }
               if (fieldList["BasementType"].ToLower().Contains("partial"))
               {
                   macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:Comparables_" + compNum + "__Structure_Basement_IsPartial&&VALUE:true CONTENT=YES");
               }
               if (fieldList["BasementType"].ToLower().Contains("walkout"))
               {
                   macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:Comparables_" + compNum + "__Structure_Basement_IsWalkout&&VALUE:true CONTENT=YES");
               }
               if (fieldList["BasementType"].ToLower().Contains("slab"))
               {
                   macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:Comparables_" + compNum + "__Structure_Basement_IsSlab&&VALUE:true CONTENT=YES");
               }
               if (fieldList["BasementType"].ToLower().Contains("crawl"))
               {
                   macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:Comparables_" + compNum + "__Structure_Basement_HasCrawlspace&&VALUE:true CONTENT=YES");
               }
               if (!string.IsNullOrEmpty(fieldList["Fireplace"]))
               {
                   macro.AppendLine(@"TAG POS=1 TYPE=INPUT:CHECKBOX FORM=ID:form0 ATTR=ID:Comparables_" + compNum + "__Structure_Amenities_HasFireplace&&VALUE:true CONTENT=YES");
               }



               //
               //Defaults
               //    
               macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:Comparables_" + compNum + "__Ratings_GeneralConditionRatingType&&VALUE:Average CONTENT=YES");
               macro.AppendLine(@"TAG POS=1 TYPE=INPUT:RADIO FORM=ID:form0 ATTR=ID:Comparables_" + compNum + "__Ratings_ConstructionQualityType&&VALUE:Average CONTENT=YES");
               macro.AppendLine(@"TAG POS=1 TYPE=SELECT FORM=ID:form0 ATTR=ID:Comparables_" + compNum + "__Site_PrimaryLocationFactorType CONTENT=%1");
               macro.AppendLine(@"TAG POS=1 TYPE=TEXTAREA FORM=ID:form0 ATTR=ID:Comparables_" + compNum + "__GeneralComments CONTENT=comments");
           }
            
           

            macro.AppendLine(@"TAG POS=1 TYPE=INPUT:SUBMIT ATTR=ID:TopSaveButton&&VALUE:SAVE<SP>THIS<SP>PAGE");


            string macroCode = macro.ToString();
             iim.iimPlayCode(macroCode, 60);
        }
    }
}
