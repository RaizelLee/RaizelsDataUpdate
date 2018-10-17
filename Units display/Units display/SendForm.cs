/******************************************************
                   Simple MAPI.NET
		      netmaster@swissonline.ch
*******************************************************/

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Win32Mapi;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Reflection;
using LiveSwitch.TextControl;


namespace Units_display
{
    /// <summary>
    /// Summary description for SendForm.
    /// </summary>
    public partial class SendForm : Form
    {
        private System.Windows.Forms.Button buttonAddrTO;
        private System.Windows.Forms.TextBox textTO;
        private System.Windows.Forms.Button buttonAddrCC;
        private System.Windows.Forms.TextBox textCC;
        private System.Windows.Forms.Label labelSubj;
        private System.Windows.Forms.TextBox textSubject;
        private System.Windows.Forms.Button buttonAttach;
        private System.Windows.Forms.ComboBox comboAttachm;
        private System.Windows.Forms.Button buttonSend;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;
        //private Editor editor = new Editor();
        private Mapi ma;
        //private string htmlString = "";
        private String sHtml;
        Outlook.Application oApp = new Outlook.Application();
        Outlook.NameSpace oNS;
        Outlook.MailItem oMsg;

        string num1_start = "<P class=MsoNormal style=\"MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\"><B><SPAN style=\'FONT-SIZE: 13.5pt; FONT-FAMILY: \"Times New Roman\",serif; COLOR: #1f4e79; mso-fareast-font-family: \"Times New Roman\"\'>Hi All,</br>The machines have arrived, please pick up if you need.</br></SPAN></B><SPAN style=\'FONT-SIZE: 13.5pt; FONT-FAMILY: \"Times New Roman\",serif; COLOR: black; mso-fareast-font-family: \"Times New Roman\"\'><o:p></o:p></SPAN></P>";
        string num1_type_open = "</br><P class=MsoNormal style=\"MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\"><B><SPAN style='FONT-SIZE: 13.5pt; FONT-FAMILY: \"Times New Roman\",serif; COLOR: #1f4e79; mso-fareast-font-family: \"Times New Roman\"\'>【";
        /// type(BNB.CNB...)
        string num1_type_close = "】<o:p></o:p></SPAN></B></P>";
        string num1_table_open = "<TABLE class=MsoNormalTable style=\"BORDER-TOP: medium none; BORDER-RIGHT: medium none; WIDTH: 211.75pt; BORDER-COLLAPSE: collapse; BORDER-BOTTOM: medium none; BORDER-LEFT: medium none; MARGIN: auto auto auto -0.5pt; mso-border-alt: solid windowtext .5pt; mso-yfti-tbllook: 1184; mso-padding-alt: 0in 0in 0in 0in; mso-border-insideh: .5pt solid windowtext; mso-border-insidev: .5pt solid windowtext\" cellSpacing=0 cellPadding=0 width=282 border=1><TBODY><TR style=\"HEIGHT: 14.3pt; mso-yfti-irow: 0; mso-yfti-firstrow: yes\"><TD style=\"BORDER-TOP: windowtext 1pt solid; HEIGHT: 14.3pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 1.75in; BACKGROUND: #00b0f0; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 5.4pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; mso-border-alt: solid windowtext .5pt\" width=168 noWrap><P class=MsoNormal style=\"MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\"><B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY: \"Microsoft JhengHei\",sans-serif; COLOR: white; mso-bidi-font-family: \"Times New Roman\"\'>Platform Name</SPAN></B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY: \"Times New Roman\",serif; mso-fareast-font-family: \"Times New Roman\"\'><o:p></o:p></SPAN></P></TD><TD style=\"BORDER-TOP: windowtext 1pt solid; HEIGHT: 14.3pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 42.25pt; BACKGROUND: #00b0f0; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 5.4pt; BORDER-LEFT: #f0f0f0; PADDING-RIGHT: 5.4pt; mso-border-alt: solid windowtext .5pt; mso-border-left-alt: solid windowtext .5pt\" width=56 noWrap><P class=MsoNormal style=\"TEXT-ALIGN: center; MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" align=center><B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY: \"Microsoft JhengHei\",sans-serif; COLOR: white; mso-bidi-font-family: \"Times New Roman\"\'>Phase</SPAN></B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY: \"Times New Roman\",serif; mso-fareast-font-family: \"Times New Roman\"\'><o:p></o:p></SPAN></P></TD><TD style=\"BORDER-TOP: windowtext 1pt solid; HEIGHT: 14.3pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 43.5pt; BACKGROUND: #00b0f0; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 5.4pt; BORDER-LEFT: #f0f0f0; PADDING-RIGHT: 5.4pt; mso-border-alt: solid windowtext .5pt; mso-border-left-alt: solid windowtext .5pt\" width=58><P class=MsoNormal style=\"TEXT-ALIGN: center; MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" align=center><B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY: \"Microsoft JhengHei\",sans-serif; COLOR: white; mso-bidi-font-family: \"Times New Roman\"\'>QTY</SPAN></B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY: \"Times New Roman\",serif; mso-fareast-font-family: \"Times New Roman\"\'><o:p></o:p></SPAN></P></TD></TR>";
        string num1_newRow_open1 = "<TR style=\"HEIGHT: 15pt; mso-yfti-irow: ";///要加是第幾行+">
        string num1_newRow_open2 = "<TD style=\"BORDER-TOP: #f0f0f0; HEIGHT: 15pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 1.75in; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 5.4pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BACKGROUND-COLOR: transparent; mso-border-alt: solid windowtext .5pt; mso-border-top-alt: solid windowtext .5pt\" width=168 noWrap><P class=MsoNormal style = \"MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" ><SPAN style='FONT-SIZE: 12pt; FONT-FAMILY: \"Microsoft JhengHei\",sans-serif; COLOR: #002060; mso-bidi-font-family: \"Times New Roman\"\'>";
        /// platform phase qty
        string num1_newRow_close = "<o:p></o:p></SPAN></P></TD>";
        // 最後一個要加</TR>
        // close table 要加 </TBODY></TABLE>
        string num1_endMail = "<P class=MsoNormal style=\"MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\"><B><SPAN style=\'FONT-SIZE: 13.5pt; FONT-FAMILY: \"Times New Roman\",serif; COLOR: #1f4e79; mso-fareast-font-family: \"Times New Roman\"\'></br>Thanks,</br>BR.</SPAN></B></P>";
        
        string num2_start = "<P class=MsoNormal style=\"MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\"><B><SPAN style=\'FONT-SIZE: 13.5pt; FONT-FAMILY: \"Times New Roman\",serif; COLOR: #1f4e79; mso-fareast-font-family: \"Times New Roman\"\'>Hi All,</br>The following is the update record of IUR, thanks.</br></SPAN></B><SPAN style=\'FONT-SIZE: 13.5pt; FONT-FAMILY: \"Times New Roman\",serif; COLOR: black; mso-fareast-font-family: \"Times New Roman\"\'><o:p></o:p></SPAN></P>";
        //string num2_type_open = "</br><P class=MsoNormal style=\"MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\"><B><SPAN style='FONT-SIZE: 13.5pt; FONT-FAMILY: \"Times New Roman\",serif; COLOR: #1f4e79; mso-fareast-font-family: \"Times New Roman\"\'>【";
        //string num2_type_close ="】<o:p></o:p></SPAN></B></P>";
        string num2_table_open = "<TABLE class=MsoNormalTable style=\"BORDER-TOP: medium none; BORDER-RIGHT: medium none; WIDTH: 954.1pt; BORDER-COLLAPSE: collapse; BORDER-BOTTOM: medium none; BORDER-LEFT: medium none; MARGIN: auto auto auto -0.5pt; mso-border-alt: solid windowtext .5pt; mso-yfti-tbllook: 1184; mso-padding-alt: 0in 0in 0in 0in; mso-border-insideh: .5pt solid windowtext; mso-border-insidev: .5pt solid windowtext\" cellSpacing=0 cellPadding=0 width=1300 border=1><TBODY><TR style=\"HEIGHT: 14.3pt; mso-yfti-irow: 0; mso-yfti-firstrow: yes\"><TD style=\"BORDER-TOP: windowtext 1pt solid; HEIGHT: 14.3pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 117pt; BACKGROUND: #00b0f0; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 5.4pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; mso-border-alt: solid windowtext .5pt\" width=130 noWrap><P class=MsoNormal style=\"TEXT-ALIGN: center; MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" align=center><B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY:\"Microsoft JhengHei\",sans-serif; COLOR: white; mso-bidi-font-family: \"Times New Roman\"\'>Platform</SPAN></B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY: \"Times New Roman\",serif; mso-fareast-font-family: \"Times New Roman\"\'><?xml:namespace prefix =\"o\" /><o:p></o:p></SPAN></P></TD><TD style=\"BORDER-TOP: windowtext 1pt solid; HEIGHT: 14.3pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 117pt; BACKGROUND: #00b0f0; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 5.4pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; mso-border-alt: solid windowtext .5pt\" width=130 noWrap><P class=MsoNormal style=\"TEXT-ALIGN: center; MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" align=center><B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY:\"Microsoft JhengHei\",sans-serif; COLOR: white; mso-bidi-font-family: \"Times New Roman\"\'>Phase</SPAN></B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY:\"Times New Roman\",serif; mso-fareast-font-family: \"Times New Roman\"\'><o:p></o:p></SPAN></P></TD><TD style=\"BORDER-TOP: windowtext 1pt solid; HEIGHT: 14.3pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 117pt; BACKGROUND: #00b0f0; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 5.4pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; mso-border-alt: solid windowtext .5pt\" width=130 noWrap><P class=MsoNormal style=\"TEXT-ALIGN: center; MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" align=center><B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY:\"Microsoft JhengHei\",sans-serif; COLOR: white; mso-bidi-font-family: \"Times New Roman\"\'>SKU</SPAN></B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY:\"Times New Roman\",serif; mso-fareast-font-family: \"Times New Roman\"\'><o:p></o:p></SPAN></P></TD><TD style=\"BORDER-TOP: windowtext 1pt solid; HEIGHT: 14.3pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 117pt; BACKGROUND: #00b0f0; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 5.4pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; mso-border-alt: solid windowtext .5pt\" width=130 noWrap><P class=MsoNormal style=\"TEXT-ALIGN: center; MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" align=center><B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY:\"Microsoft JhengHei\",sans-serif; COLOR: white; mso-bidi-font-family: \"Times New Roman\"\'>S/N<o:p></o:p></SPAN></B></P></TD><TD style=\"BORDER-TOP: windowtext 1pt solid; HEIGHT: 14.3pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 117pt; BACKGROUND: #00b0f0; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 5.4pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; mso-border-alt: solid windowtext .5pt\" width=130 noWrap><P class=MsoNormal style=\"TEXT-ALIGN: center; MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" align=center><B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY:\"Microsoft JhengHei\",sans-serif; COLOR: white; mso-bidi-font-family: \"Times New Roman\"\'>Owner<o:p></o:p></SPAN></B></P></TD><TD style=\"BORDER-TOP: windowtext 1pt solid; HEIGHT: 14.3pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 117pt; BACKGROUND: #00b0f0; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 5.4pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; mso-border-alt: solid windowtext .5pt\" width=130 noWrap><P class=MsoNormal style=\"TEXT-ALIGN: center; MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" align=center><B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY:\"Microsoft JhengHei\",sans-serif; COLOR: white; mso-bidi-font-family: \"Times New Roman\"\'>Status<o:p></o:p></SPAN></B></P></TD><TD style=\"BORDER-TOP: windowtext 1pt solid; HEIGHT: 14.3pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 117pt; BACKGROUND: #00b0f0; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 5.4pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; mso-border-alt: solid windowtext .5pt\" width=130 noWrap><P class=MsoNormal style=\"TEXT-ALIGN: center; MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" align=center><B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY:\"Microsoft JhengHei\",sans-serif; COLOR: white; mso-bidi-font-family: \"Times New Roman\"\'>Year<o:p></o:p></SPAN></B></P></TD><TD style=\"BORDER-TOP: windowtext 1pt solid; HEIGHT: 14.3pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 117pt; BACKGROUND: #00b0f0; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 5.4pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; mso-border-alt: solid windowtext .5pt\" width=130 noWrap><P class=MsoNormal style=\"TEXT-ALIGN: center; MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" align=center><B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY:\"Microsoft JhengHei\",sans-serif; COLOR: white; mso-bidi-font-family: \"Times New Roman\"\'>Category<o:p></o:p></SPAN></B></P></TD><TD style=\"BORDER-TOP: windowtext 1pt solid; HEIGHT: 14.3pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 117pt; BACKGROUND: #00b0f0; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 5.4pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; mso-border-alt: solid windowtext .5pt\" width=130 noWrap><P class=MsoNormal style=\"TEXT-ALIGN: center; MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" align=center><B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY:\"Microsoft JhengHei\",sans-serif; COLOR: white; mso-bidi-font-family: \"Times New Roman\"\'>Note<o:p></o:p></SPAN></B></P></TD><TD style=\"BORDER-TOP: windowtext 1pt solid; HEIGHT: 14.3pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 117pt; BACKGROUND: #00b0f0; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 5.4pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; mso-border-alt: solid windowtext .5pt\" width=130 noWrap><P class=MsoNormal style=\"TEXT-ALIGN: center; MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" align=center><B><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY:\"Microsoft JhengHei\",sans-serif; COLOR: white; mso-bidi-font-family: \"Times New Roman\"\'>Date<o:p></o:p></SPAN></B></P></TD></TR>";

        string num2_rowOpen1 = "<TR style=\"HEIGHT: 15pt; mso-yfti-irow:"; ///要加是第幾行+">
        string num2_rowOpen2 = "<TD style=\"BORDER-TOP: #f0f0f0; HEIGHT: 15pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 99.1pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 5.4pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BACKGROUND-COLOR: transparent; mso-border-alt: solid windowtext .5pt; mso-border-top-alt: solid windowtext .5pt\" width=132 noWrap><P class=MsoNormal style = \"TEXT-ALIGN: center; MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" align=center><SPAN style = \'FONT-SIZE: 12pt; FONT-FAMILY: \"Microsoft JhengHei\",sans-serif; COLOR: #002060; mso-bidi-font-family: \"Times New Roman\"\' > ";
        /// text here
        string num2_rowOpen2_col4 = "<TD style=\"BORDER-TOP: #f0f0f0; HEIGHT: 15pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 1.5in; BACKGROUND: #fff2cc; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 0in; BORDER-LEFT: #f0f0f0; PADDING-RIGHT: 0in; mso-border-alt: solid windowtext .5pt; mso-border-left-alt: solid windowtext .5pt; mso-border-top-alt: solid windowtext .5pt; mso-background-themecolor: accent4; mso-background-themetint: 51\" vAlign=top width=144><P class=MsoNormal style=\"TEXT-ALIGN: center; MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" align=center><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY: \"Microsoft JhengHei\",sans-serif; COLOR: #002060; mso-bidi-font-family:\"Times New Roman\"\'>";
        string num2_rowOpen2_col5 = "<TD style=\"BORDER-TOP: #f0f0f0; HEIGHT: 15pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 117pt; BACKGROUND: yellow; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0in; PADDING-TOP: 0in; PADDING-LEFT: 0in; BORDER-LEFT: #f0f0f0; PADDING-RIGHT: 0in; mso-border-alt: solid windowtext .5pt; mso-border-left-alt: solid windowtext .5pt; mso-border-top-alt: solid windowtext .5pt\" vAlign=top width=156><P class=MsoNormal style=\"TEXT-ALIGN: center; MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal\" align=center><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY: \"Microsoft JhengHei\",sans-serif; COLOR: #002060; mso-bidi-font-family: \"Times New Roman\"\'>";
        string num2_rowClose = "</SPAN><SPAN style=\'FONT-SIZE: 12pt; FONT-FAMILY: \"Times New Roman\",serif; mso-fareast-font-family: \"Times New Roman\"\'><o:p></o:p></SPAN></P></TD>";
        /// 一行最後要加</TR>
        /// table close要加</TBODY></TABLE></SPAN></P>
        //private Font boldFont;
        //MailEnvelop currentMail;
        //MailComparer comparer = new MailComparer();
        public Form3 preForm;
        string[] ph = { };
        private Form1.unitTB[] unitsTB;
        private Form1.unitInfo[] unitInfos;

        public SendForm(Form callingForm, ref Mapi rma)
        {
            ma = rma;
            oNS = oApp.GetNamespace("mapi");
            oNS.Logon(Missing.Value, Missing.Value, true, true);
            oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
            preForm = callingForm as Form3;
            ma.Reset();
            InitializeComponent();
            Form1.unitTB[] unitsTB = preForm.returnUnitTB();
            Form1.unitInfo[] unitInfos = preForm.returnUnitInfo();

            if (preForm.returnChoose() == 1)
            {
                //textTO.Text = "Lee, Simon <simon.lee@hp.com>; Hsu, Aurora <aurora.hsu@hp.com>; HUNG, SEAN (Sean Hung) <sean.hung@hp.com>; Tsai, Hunter <hunter.tsai@hp.com>; Hsieh, David (David Hsieh) <david.hsieh@hp.com>; Chou, Cindy <cindy.chou@hp.com>; Cheng, Steven <steven.cheng2@hp.com>; Su, Demi <demis@hp.com>; Chen, Frank (COMMs TDC) <frank.chen@hp.com>; Su, YJ <yj.su@hp.com>; Lee, Connant <connant.lee@hp.com>; Tseng, YJ <disabled-yj.tseng@hp.com>; Wu, Jason <jason.wu@hp.com>; Lin, Robin ((COMM TDC)) <robin.lin@hp.com>; Yang, Daniel (Comm TDC) <daniel.yang@hp.com>; Chen, Pat <pat.chen@hp.com>; Lin, CF <cf.lin@hp.com>; Chen, Troy (CMIT NB RandD) <troyc@hp.com>; Hung, KC <kc.hung@hp.com>; Lu, Jacky <jacky.lu@hp.com>; Chiang, Bernice <bernice.chiang@hp.com>; Chan, Tony (WWAN) <tony.chan2@hp.com>; Lee, Amanda <amanda.lee@hp.com>; Chan, Tony (WWAN) <tony.chan2@hp.com>; Chien, Jeremy (Comms Radio HW) <jeremy.chien@hp.com>; Chou, Emily (CMIT) <emily.chou@hp.com>; Wu, Kent <kent.wu@hp.com>; Chen, Xinchang (Comm TDC) <xinchang.chen@hp.com>; Ma, Terry <terry.ma@hp.com>; Lin, Steve <steve.lin@hp.com>; Wu, Anthony <anthony.wu@hp.com>; Chou, Emily (CMIT) <emily.chou@hp.com>; Cheng, Timothy <timothy.cheng@hp.com>; Liu, Bruce <brucel@hp.com>; Huang, Jake <jake.huang@hp.com>; Tung, Grace <grace.tung@hp.com>; Kuo, Jason <jason.kuo2@hp.com>; You, Ethan (Comm TDC) <ethan.you@hp.com>; Zeng, Grey <grey.zeng@hp.com>; Fong, Agness <agness.fong1@hp.com>; Liu, Sean (Comm TDC) <seanl@hp.com>; Yu, Debby <debby.yu@hp.com>; Hsieh, Sean <sean.hsieh@hp.com>; Hsieh, James <james.hsieh@hp.com>; Lin, Tony <tony.lin@hp.com>; Tsai, David <david.tsai@hp.com>; Cheng, Mike <mike.cheng@hp.com>; Chu, Jimmy <jimmy.chu1@hp.com>";
                //textCC.Text = "Chang, Hans <hans.chang@hp.com>; Hung, KC <kc.hung@hp.com>; Cheng, Steven <steven.cheng2@hp.com>";
                textSubject.Text = "Machines arrived ";
                Array.Clear(ph,0,ph.Length);
                //if (ph.Length != 0)
                //{
                //    Array.Resize(ref ph, 0);
                //}
                bool phExist = false;
                for(int j=0; j<unitsTB.Length; j++)
                {
                    for (int i = 0; i < ph.Length; i++)
                    {
                        if(unitsTB[j].platform == ph[i])
                        {
                            phExist = true;
                        }
                    }
                    if (phExist == false)
                    {
                        Array.Resize(ref ph, ph.Length + 1);
                        ph[ph.Length - 1] = unitsTB[j].platform;
                    }
                    phExist = false;
                }
                for(int i=0; i< ph.Length; i++)
                {
                    if(ph[i]!=null)
                        textSubject.Text += "【" + ph[i] + "】";
                }

                sHtml += num1_start;
                createTable("CNB");
                createTable("BNB");
                createTable("CDT");
                createTable("BDT");
                createTable("CAIO");
                createTable("BAIO");
                sHtml += num1_endMail;
            }
            else
            {
                textSubject.Text = "IUR record update";
                sHtml = num2_start+ num1_type_open;
                if (unitInfos[0].Borrower =="Storage"|| unitInfos[0].Borrower == "storage")
                {
                    sHtml += "Return list";
                }
                else{
                    sHtml += "Borrow list";
                }
                sHtml += num1_type_close + num2_table_open;
                for (int i=0; i<unitInfos.Length; i++)
                {
                    sHtml += num2_rowOpen1 + (i + 1).ToString() + "\">" +
                        num2_rowOpen2 + unitInfos[i].Platform.ToString() + num2_rowClose +
                        num2_rowOpen2 + unitInfos[i].Phase.ToString() + num2_rowClose +
                        num2_rowOpen2 + unitInfos[i].SKU.ToString() + num2_rowClose +
                        num2_rowOpen2_col4 + unitInfos[i].SN.ToString() + num2_rowClose +
                        num2_rowOpen2_col5 + unitInfos[i].Borrower.ToString() + num2_rowClose +
                        num2_rowOpen2 + unitInfos[i].Status.ToString() + num2_rowClose +
                        num2_rowOpen2 + unitInfos[i].Year.ToString() + num2_rowClose +
                        num2_rowOpen2 + unitInfos[i].Category.ToString() + num2_rowClose +
                        num2_rowOpen2 + unitInfos[i].Note.ToString() + num2_rowClose +
                        num2_rowOpen2 + unitInfos[i].Date.Substring(0,10).ToString() + num2_rowClose + "</TR>";
                }
                sHtml += "</TBODY></TABLE></SPAN></P>" + num1_endMail;
                //string htmltext = editor1.Html.ToString();
                //System.IO.File.WriteAllText(@"C:\Users\HP\Desktop\htmltext.txt", htmltext);
            }
            editor1.Html = sHtml;
            //System.IO.File.WriteAllText(@"C:\Users\HP\Desktop\htmltext.txt", editor1.Html);

        }

        private void createTable(string type)
        {
            bool typeTitle = false;
            int rowCount = 0;
            unitsTB = preForm.returnUnitTB();

            for (int i = 0; i < unitsTB.Length; i++)
            {
                if (unitsTB[i].type == type)
                {
                    if (typeTitle == false)
                    {
                        typeTitle = true;
                        sHtml += num1_type_open + type + num1_type_close + num1_table_open;
                    }
                    rowCount++;
                    sHtml += num1_newRow_open1 + rowCount.ToString() + "\">" +
                        num1_newRow_open2 + unitsTB[i].platform + num1_newRow_close +
                        num1_newRow_open2 + unitsTB[i].phase + num1_newRow_close +
                        num1_newRow_open2 + unitsTB[i].qty.ToString() + num1_newRow_close;
                }
            }
            if(typeTitle == true)
            {
                sHtml += "</TR></TBODY></TABLE>" + num1_newRow_close;
            }
        }
        public SendForm(ref Mapi rma)
        {
            //editor.FormBorderStyle = FormBorderStyle.None; // 无边框
            //editor.TopLevel = false; // 不是最顶层窗体
            //panel1.Controls.Add(editor);  // 添加到 Panel中
            //editor.Show();

        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>

        private void editor_Tick()
        {
            //undoToolStripMenuItem.Enabled = editor.CanUndo();
            //redoToolStripMenuItem.Enabled = editor.CanRedo();
            //cutToolStripMenuItem.Enabled = editor.CanCut();
            //copyToolStripMenuItem.Enabled = editor.CanCopy();
            //pasteToolStripMenuItem.Enabled = editor.CanPaste();
            //imageToolStripMenuItem.Enabled = editor.CanInsertLink();
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SendForm));
            this.textTO = new System.Windows.Forms.TextBox();
            this.textSubject = new System.Windows.Forms.TextBox();
            this.labelSubj = new System.Windows.Forms.Label();
            this.textCC = new System.Windows.Forms.TextBox();
            this.buttonSend = new System.Windows.Forms.Button();
            this.comboAttachm = new System.Windows.Forms.ComboBox();
            this.buttonAttach = new System.Windows.Forms.Button();
            this.buttonAddrTO = new System.Windows.Forms.Button();
            this.buttonAddrCC = new System.Windows.Forms.Button();
            this.editor1 = new LiveSwitch.TextControl.Editor();
            this.SuspendLayout();
            // 
            // textTO
            // 
            this.textTO.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textTO.Location = new System.Drawing.Point(64, 8);
            this.textTO.MaxLength = 5000;
            this.textTO.Name = "textTO";
            this.textTO.Size = new System.Drawing.Size(644, 20);
            this.textTO.TabIndex = 2;
            // 
            // textSubject
            // 
            this.textSubject.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textSubject.Location = new System.Drawing.Point(64, 104);
            this.textSubject.MaxLength = 300;
            this.textSubject.Name = "textSubject";
            this.textSubject.Size = new System.Drawing.Size(644, 20);
            this.textSubject.TabIndex = 3;
            // 
            // labelSubj
            // 
            this.labelSubj.Location = new System.Drawing.Point(7, 107);
            this.labelSubj.Name = "labelSubj";
            this.labelSubj.Size = new System.Drawing.Size(49, 16);
            this.labelSubj.TabIndex = 1;
            this.labelSubj.Text = "Subject";
            // 
            // textCC
            // 
            this.textCC.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textCC.Location = new System.Drawing.Point(64, 40);
            this.textCC.MaxLength = 5000;
            this.textCC.Name = "textCC";
            this.textCC.Size = new System.Drawing.Size(644, 20);
            this.textCC.TabIndex = 7;
            // 
            // buttonSend
            // 
            this.buttonSend.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonSend.Location = new System.Drawing.Point(624, 460);
            this.buttonSend.Name = "buttonSend";
            this.buttonSend.Size = new System.Drawing.Size(80, 24);
            this.buttonSend.TabIndex = 5;
            this.buttonSend.Text = "Send!";
            this.buttonSend.Click += new System.EventHandler(this.buttonSend_Click);
            // 
            // comboAttachm
            // 
            this.comboAttachm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboAttachm.DropDownWidth = 300;
            this.comboAttachm.Enabled = false;
            this.comboAttachm.Location = new System.Drawing.Point(64, 72);
            this.comboAttachm.Name = "comboAttachm";
            this.comboAttachm.Size = new System.Drawing.Size(644, 21);
            this.comboAttachm.TabIndex = 9;
            // 
            // buttonAttach
            // 
            this.buttonAttach.Location = new System.Drawing.Point(8, 71);
            this.buttonAttach.Name = "buttonAttach";
            this.buttonAttach.Size = new System.Drawing.Size(56, 24);
            this.buttonAttach.TabIndex = 8;
            this.buttonAttach.Text = "Attach...";
            this.buttonAttach.Click += new System.EventHandler(this.buttonAttach_Click);
            // 
            // buttonAddrTO
            // 
            this.buttonAddrTO.Location = new System.Drawing.Point(8, 5);
            this.buttonAddrTO.Name = "buttonAddrTO";
            this.buttonAddrTO.Size = new System.Drawing.Size(56, 24);
            this.buttonAddrTO.TabIndex = 1;
            this.buttonAddrTO.Text = "TO...";
            this.buttonAddrTO.Click += new System.EventHandler(this.buttonAddrTO_Click);
            // 
            // buttonAddrCC
            // 
            this.buttonAddrCC.Location = new System.Drawing.Point(8, 38);
            this.buttonAddrCC.Name = "buttonAddrCC";
            this.buttonAddrCC.Size = new System.Drawing.Size(56, 24);
            this.buttonAddrCC.TabIndex = 6;
            this.buttonAddrCC.Text = "CC...";
            this.buttonAddrCC.Click += new System.EventHandler(this.buttonAddrCC_Click);
            // 
            // editor1
            // 
            this.editor1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.editor1.BodyBackgroundColor = System.Drawing.Color.White;
            this.editor1.BodyHtml = null;
            this.editor1.BodyText = null;
            this.editor1.DocumentText = resources.GetString("editor1.DocumentText");
            this.editor1.EditorBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.editor1.EditorForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.editor1.FontSize = LiveSwitch.TextControl.FontSize.Three;
            this.editor1.Html = null;
            this.editor1.Location = new System.Drawing.Point(8, 143);
            this.editor1.Name = "editor1";
            this.editor1.Size = new System.Drawing.Size(696, 307);
            this.editor1.TabIndex = 11;
            // 
            // SendForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(716, 496);
            this.Controls.Add(this.editor1);
            this.Controls.Add(this.comboAttachm);
            this.Controls.Add(this.buttonSend);
            this.Controls.Add(this.textTO);
            this.Controls.Add(this.textSubject);
            this.Controls.Add(this.labelSubj);
            this.Controls.Add(this.textCC);
            this.Controls.Add(this.buttonAttach);
            this.Controls.Add(this.buttonAddrTO);
            this.Controls.Add(this.buttonAddrCC);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(400, 280);
            this.Name = "SendForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Send Mail";
            this.Activated += new System.EventHandler(this.SendForm_Activated);
            this.Load += new System.EventHandler(this.SendForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion
        int sendNum = 1;
        private void buttonSend_Click(object sender, System.EventArgs e)
        {

            //string htmltext = editor1.Html.ToString();
            //System.IO.File.WriteAllText(@"C:\Users\HP\Desktop\htmltext.txt", htmltext);

            if (sendNum != 0)
            {
                if ((textTO.Text == null) || (textSubject.Text == null))
                    return;
                if ((textTO.Text.Length == 0) || (textSubject.Text.Length == 0))
                    return;

                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(textTO.Text);

                if (textCC.Text != null)
                {
                    if (textCC.Text.Length > 0)
                    {
                        Outlook.Recipients ccRecips = (Outlook.Recipients)oMsg.Recipients;
                        Outlook.Recipient ccRecip = (Outlook.Recipient)ccRecips.Add(textTO.Text);
                        ccRecip.Type = (int)Outlook.OlMailRecipientType.olCC;
                    }
                    //ma.AddRecip(textCC.Text, null, true);
                }

                oMsg.Subject = textSubject.Text;

                oMsg.HTMLBody = editor1.Html;
                //if (!ma.Send(textSubject.Text, textMail.Text))
                //    MessageBox.Show(this, "MAPISendMail failed! " + ma.Error(), "Send Mail", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                oRecip.Resolve();
                // Send.
                oMsg.Send();

                // Log off.
                oNS.Logoff();

                // Clean up.
                oRecip = null;
                oRecips = null;
                oMsg = null;
                oNS = null;
                oApp = null;

                ma.Reset();
                this.Close();
            }
        }

        private void buttonAttach_Click(object sender, System.EventArgs e)
        {
            OpenFileDialog af = new OpenFileDialog();
            af.Multiselect = true;
            af.Title = "Attach File";
            af.Filter = "Any File (*.*)|*.*";

            if (af.ShowDialog() != DialogResult.OK)
                return;

            comboAttachm.Enabled = true;
            int n = 0;
            //comboAttachm.Items.Add(af.FileName);
            //comboAttachm.SelectedIndex = n;
            foreach (string strFilename in af.FileNames)
            {
                oMsg.Attachments.Add(af.FileNames[n],
                Outlook.OlAttachmentType.olByValue, Type.Missing,
                Type.Missing);
                comboAttachm.Items.Add(af.FileNames[n]);
                n++;
            }
            comboAttachm.SelectedIndex = n-1;
        }

        private void buttonAddrTO_Click(object sender, System.EventArgs e)
        {
            string name; string addr;
            if (ma.SingleAddress(null, out name, out addr))
                textTO.Text = name;
        }

        private void buttonAddrCC_Click(object sender, System.EventArgs e)
        {
            string name; string addr;
            if (ma.SingleAddress("CC", out name, out addr))
                textCC.Text = name;
        }

        private void SendForm_Load(object sender, EventArgs e)
        {

            //editor1.Html = System.IO.File.ReadAllText(@"C:\Users\HP\Desktop\htmltext.txt");
            //wb.DocumentText = sHtml;
        }

        private void SendForm_Activated(object sender, EventArgs e)
        {
        //    if (!first_activated)
        //    {
        //        first_activated = true;
        //        ma.Logon(this.Handle);
        //            //RefreshInbox();
        //    }
        }

        private void textMail_TextChanged(object sender, EventArgs e)
        {

        }

        //private void button1_Click(object sender, EventArgs e)
        //{
        //    //string htmltext = editor1.Html.ToString();
        //    System.IO.File.WriteAllText(@"C:\Users\HP\Desktop\htmltext.txt", editor1.Html);
        //}

        
    }
}
