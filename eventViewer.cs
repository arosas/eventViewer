using System;
using System.Web;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Serialization;
using System.Text;
using System.Collections.Generic;
using System.Web.UI.HtmlControls;
using System.Text.RegularExpressions;
 
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint.WebPartPages;
using Microsoft.SharePoint.Utilities;
using System.Net.Mail;
 
namespace EventViewer
{
    [Guid("53f390d9-2131-46ea-8d12-75d299ac012a")]
    public class EventViewer : System.Web.UI.WebControls.WebParts.WebPart
    {
        public EventViewer()
        {
            this.ExportMode = WebPartExportMode.All;
        }
        /************************************
         *      declared variables          *
         * -------------------------------- *
         * | dropdown menu    - _dropdown | *
         * | Year             - _year     | *
         * | item limit       - _itemLimit| *
         * | list name        - _listName | *
         * | upcoming events  - _upcoming | *
         * | web              - _web      | *
         * -------------------------------  *
         * **********************************/
        //private Microsoft.SharePoint.SPWeb _web = null;
        //private Microsoft.SharePoint.SPList _eventList = null;
        private ScrollDirection _dropdown;
        private string _year, _itemLimit, _listName, _listURL;
        private bool _upcoming;
 
        /**
         *  GetListByName - private method to get any list by either
         *                  using a text box or a drop down menu.
         **/
        //private SPList GetListByName(string _listName)
        //{
        //    using (SPWeb oSPWeb = SPControl.GetContextWeb(HttpContext.Current))
        //    {
        //        SPListCollection lists = oSPWeb.Lists;
        //        return lists[_listName];
        //    }
        //}
 
        #region event viewer webpart control properties
 
        #region list textbox - get the list url requested by the user
 
        [WebBrowsable(true),
        Category("Display Properties"),
        WebDisplayName("List Name"),
        WebDescription("List Name"),
        Personalizable(PersonalizationScope.Shared),
        XmlElement(ElementName = "listName")
        ]
        public string listName{
            get { return _listName; }
            set { _listName = value; }
        }
 
        #endregion
 
        #region list textbox - get the list requested by the user
 
        [WebBrowsable(true),
        Category("Display Properties"),
        WebDisplayName("List URL"),
        WebDescription("List URL"),
        Personalizable(PersonalizationScope.Shared),
        XmlElement(ElementName = "listURL")]
        
        public string listURL{
            get { return _listURL; }
            set { _listURL = value; }
        }
 
        #endregion
 
        #region year textbox - this is where the user enters the year
 
        [WebBrowsable(true),
        Category("Display Properties"),
        WebDisplayName("Event Year"),
        WebDescription("Event Year"),
        Personalizable(PersonalizationScope.Shared),
        XmlElement(ElementName = "Year")]
        public string EventYear{
            get { return _year; }
            set { _year = value; }
        }
 
        #endregion
 
        #region upcoming events checkbox - user checks to see only upcomming events
 
        #region upcomming events only
 
        [WebBrowsable(true),
        Category("Display Properties"),
        WebDisplayName("Upcoming Events"),
        WebDescription("Upcoming Events"),
        Personalizable(PersonalizationScope.Shared),
        XmlElement(ElementName = "Upcoming")]
 
        public bool upcomming{
            get { return _upcoming; }
            set { _upcoming = value; }
        }
 
        #endregion
 
        #endregion
 
        #region dropdown menu has the long and short view options, so the user can select one of the 2 options
        public enum ScrollDirection{
            Long, Short
        };
        [WebBrowsable(true),
       Category("Display Properties"),
        DefaultValue("2008"),
        WebDisplayName("View"),
        WebDescription("select the desired view"),
        Personalizable(PersonalizationScope.Shared),
        XmlElement(ElementName = "View Details")]

        public ScrollDirection DdDirection{
            get{
                return this._dropdown;
            }
            set{
                this._dropdown = value;
            }
        }
        #endregion
 
        #region item limit text box - user sets the limit to the items listed on the textbox
        int counter = 0;
 
        [WebBrowsable(true),
       Category("Display Properties"),
        WebDisplayName("Item Limit"),
        WebDescription("Item Limit"),
        Personalizable(PersonalizationScope.Shared),
        XmlElement(ElementName = "ItemLimit")]

        public string itemLimit{
            get { return _itemLimit; }
            set { _itemLimit = value; }
        }
        #endregion
 
        #endregion
 
        //webpart renderer
        protected override void Render(HtmlTextWriter writer){
            
            if (_listName == null){
                writer.Write("<p style=\"text-align:center; margin-top:15px\"><span style=\"font-size:12px;"
                + "font-weight:bold; color:Red\">Please set List Name in the Display Properties section of "
                + "the web part</br></br></span></p>");
            }
            else{

                string _eventsSite = this._listURL;
 
                using (SPSite site = new SPSite(_eventsSite)){
                    using (SPWeb _web = site.OpenWeb())
                    {
                        SPList _eventList = _web.Lists[this._listName];
 
                        // give this a try!
                        try{
                            //get to the web
                            //this._web = Microsoft.SharePoint.SPContext.Current.Web;
 
                            //gets the list
 
                            //_eventList = this._web.Lists[this._listName];
                            // _eventList = _web.GetListFromUrl(_listName);
 
 
                            //builds the query by using a string builder
                            StringBuilder qString = new StringBuilder();
                            qString.Append("<Where>");
 
                            // if this years matches the _upcomming is true
                            // then it only gives out the upcomming events only if the checkbox is checked.
                            if (_year == DateTime.Now.Year.ToString() && _upcoming == true)
                            {
                                qString.Append("<Gt>");
                                qString.Append("<FieldRef Name='Expires' />");
                                qString.Append("<Value Type='DateTime'>" + DateTime.Today.Year + "-" + DateTime.Today.Month + "-" + DateTime.Today.Day + "T12:00:00Z</Value>");
                                qString.Append("</Gt>");
                            }
                            // otherwise give me everything!
                            else{
                                qString.Append("<And>");
                                qString.Append("<Geq>");
                                qString.Append("<FieldRef Name='Expires' />");
                                qString.Append("<Value Type='DateTime'>" + _year + "-01-01T12:00:00Z</Value>");
                                qString.Append("</Geq>");
                                qString.Append("<Leq>");
                                qString.Append("<FieldRef Name='Expires' />");
 
                                // if the year you entered in the textbox is the same as this year
                                // then show the dates from here on out. BAM! ^o^
                                if (_year == DateTime.Now.Year.ToString()){
                                    qString.Append("<Value Type='DateTime'>" + DateTime.Today.Year + "-" + DateTime.Today.Month + "-" + DateTime.Today.Day + "T12:00:00Z</Value>");
                                }
                                // otherwise give out the events from that year to the future! ^o^
                                else{
                                    qString.Append("<Value Type='DateTime'>" + _year + "-12-31T12:00:00Z</Value>");
                                }
                                //this appends something I think its from the caml query code.... o_O
                                qString.Append("</Leq>");
                                qString.Append("</And>");
                            }
                            //sorts the list entered in decending order if its false, true if its done in accending order
                            qString.Append("</Where>");
                            qString.Append("<OrderBy>");
 
                            if (_upcoming == true){
                                qString.Append("<FieldRef Name='Expires' Ascending='True' />");
 
                            }
                            else{
                                qString.Append("<FieldRef Name='Expires' Ascending='False' />");
                            }
                            qString.Append("</OrderBy>");
 
                            // make the query called q ^o^
                            SPQuery q = new SPQuery();
                            // put the list into the query! ^o^
                            q.Query = qString.ToString();
 
                            // EventlistItems - gets everything from the query onto the SPListItemCollection
                            SPListItemCollection EventlistItems = _eventList.GetItems(q);
 
                            // if the list is empty or -1 items which is impossible to get,
                            // then tell me that I don't have any events at all even if the
                            // upcomming checkbox is checked or not ^o^
                            if (EventlistItems.Count < 1){
                                writer.Write("No Events Scheduled for this time period");
                            }
                            // otherwise gimme the events! >_<#
                            else{
                                // this is a table tag. we use them to make tables in html. ^o^
                                writer.Write("<table>");
 
                                // if there is no limit set to the list for the given year
                                // then set the item limit to zero and keep going until the
                                // list runs out of events!!! ^o^
                                if (itemLimit == String.Empty){
                                    itemLimit = "0";
                                }

                                // enter the list
                                foreach (SPItem ItemEvent in EventlistItems){
                                    //removing the description from the link
                                    string eventUrl = ItemEvent["Event URL"].ToString();
                                    //int e = eventUrl.IndexOf(',');
                                    //eventUrl = eventUrl.Substring(0, e);
 
                                    // while its counting down to my limit
                                    if (counter < Convert.ToInt32(itemLimit) || Convert.ToInt32(itemLimit) == 0){
                                        // make a new row for my table
                                        // Did I mentioned that I made a table?
                                        writer.Write("<tr>");
 
                                        // if the user selected the long view
                                        // note: the long view is actually zero just a heads up! ^_^
                                        if (_dropdown == 0){
                                            // the value imgUrl is created!
                                            string imgUrl;
 
                                            // if the event listed does not have an image
                                            if (ItemEvent["Event Image"] == null){
                                                // input a generic picture next to the event
                                                imgUrl = "/_layouts/images/events.gif";
                                            }
                                            else{
                                                // otherwise put the picture there given with the event
                                                //
                                                // note: sharepoint is being a b**** so I have to take off
                                                //        the the description by hand
                                                imgUrl = ItemEvent["Event Image"].ToString();
                                                int n = imgUrl.IndexOf(',');
                                                imgUrl = imgUrl.Substring(0, n);
 
                                                //event url
                                                // do same as img url
                                            }
                                            // this renders the picture
                                            writer.Write("<td style='vertical-align:middle; text-align:center;'>");
                                            writer.Write("<img src='" + imgUrl + "'/>");
                                            writer.Write("</td>");
                                        }
                                        //renders the event
                                        writer.Write("<td style='vertical-align:top; text-align:left;'>");
                                        writer.Write("<p>");
                                        //writer.Write("<a href='" + ItemEvent["Event URL"] + "' target=_blank>" + ItemEvent["Title"] + "</a>");
                                        writer.Write("<a href='" + eventUrl + "' target=_blank>" + ItemEvent["Title"] + "</a>");
                                        writer.Write("<br>");
                                        writer.Write(ItemEvent["Event Date"]);
                                        writer.Write("<br>");
 
                                        // if I am in long view
                                        // note: Did I mentioned that long view is zero? O_o?
                                        if (_dropdown == 0){
                                            // if the venue is not empty.
                                            if (ItemEvent["Venue"] != null){
                                                // then render the venue ^o^
                                                writer.Write(ItemEvent["Venue"]);
                                                writer.Write("<br>");
                                            }
 
                                        }
                                        // renders the event location
                                        writer.Write(ItemEvent["Event Location"]);
                                        // paragraph
                                        writer.Write("</p>");
                                        // make an empty paragraph because Internet Explorer is being stubborn as a mule!
                                        writer.Write("<p></p></td></tr>");
 
                                        // increment by +1
                                        counter++;
                                    }
                                }
                                //close html table tab ^o^
                                writer.Write("</table>");
                            }
                        }
                        // In case I caught an error
                        catch (Exception ex){
                            // tell me what went wrong so I can fix either fix my code
                            // or tell you that you screwed up by putting an invalid option
                            notifyAdministrator("Web Part Error: Event Viewer", ex);
                            writer.Write("There was an error loading the Events.  Please try again later.</br></br>Thank you.");
                        }
                    }
                }
            }
        }
 
        private void notifyAdministrator(string subject, Exception error){
            MailMessage mailMessage = new MailMessage("no-reply@", "admin@arosas.me");
            mailMessage.IsBodyHtml = true;
            mailMessage.Priority = MailPriority.High;
            mailMessage.Subject = subject;
            StringBuilder body = new StringBuilder();
            body.Append("<head><META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=us-ascii'>");
            body.Append("<style>BODY {font-family:Verdana; font-size:9.0pt; color:#4f4c4d} </style></head>");
            body.Append("The following error has occurred while loading the Event Viewer:<br /><br />");
            body.Append("Message:<br /><br />");
            body.Append(error.ToString());
            body.Append("Stack Trace:<br /><br />");
            body.Append(error.StackTrace);
            mailMessage.Body = body.ToString();
            SmtpClient client = new SmtpClient("mail.arosas.me");
            //SmtpClient client = new SmtpClient("Localhost");
            client.Send(mailMessage);
        }
    }
}