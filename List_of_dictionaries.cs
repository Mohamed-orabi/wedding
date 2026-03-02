// Type: Issues_Application.List_of_dictionaries
// Assembly: Issues Application, Version=1.0.4828.18025, Culture=neutral, PublicKeyToken=null
// MVID: 51ADFE54-690F-445C-8399-1F6BEA2DE42C
// Assembly location: D:\Spaces\ESAMI\applications\Copy of Issues Application\bin\Debug\Issues Application.exe

//using Word;
//using Excel;
//using Microsoft.Office.Interop.Word;
//using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Oracle.DataAccess.Client;                //using System.Data.OracleClient;
using System;
using System.ComponentModel;
using System.Diagnostics.Contracts;
using System.Drawing;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Issues_Application
{
    public class List_of_dictionaries : Form
    {
        private IContainer components;
        private CheckedListBox checkedListBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button button4;
        private LinkLabel linkLabel1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private CheckedListBox checkedListBox2;
        private System.Windows.Forms.Label label2;
        private LinkLabel linkLabel2;
        private System.Windows.Forms.CheckBox checkBox1;
        private string contract_type;
        private string BaseInstallmentSettings;
        private string BaseInstallment;
        private string LinkedContractTypes;
        private string contract_status;
        private string Account_type;
        private string card_financial_profiles;
        private string Retailers_financial_profiles;
        private string Telebank_financial_profiles;
        private string SMS_financial_profiles;
        private string SMS_Channels_financial_profiles;
        private string stat_lang;
        private string limit_groups;
        private string usage_limits;
        private string countries;
        private string currencies;
        private string regions;
        private string cities;
        private string min_payments;
        private string calculation_profiles;
        private string direct_debit_profile;
        private string occupation;
        private string branches;
        private string card_status;
        private string card_state;
        private string marital_status;
        private string interest_rate;
        private string card_product;
        private string titles;
        private string pan_ranges;
        private string bank_operators;
        private string mscc_operators;
        private string Dict_Schema;
        private string Dict_Channel;
        private string Education;
        private string country_groups;
        private string mcc_groups;
        private string mcc;
        private string card_USA;
        private string account_USA;
        private string contract_USA;
        private string client_USA;
        private string sms_dic; // islam atta

        public static Word.Application oWord;
        public static _Document oDoc;

        public static object oEndOfDoc;
        public static object filename;
        public static Microsoft.Office.Interop.Excel.Application m_objExcel;
        public static Workbooks m_objBooks;
        public static _Workbook m_objBook;
        public static Sheets m_objSheets;
        public static _Worksheet m_objSheet;
        public static Microsoft.Office.Interop.Excel.Range m_objRange;
        public static Microsoft.Office.Interop.Excel.Font m_objFont;
        public static Microsoft.Office.Interop.Excel.Style m_objstyle;
        public static object m_objOpt;
        public static object col_width;
        public static object direction;
        public static object format;
        public static object tab_direction;
        public static object auto_fit;
        public static object auto_border;
        public static object auto_heading;
        public static object apply_color;
        public static object apply_header;
        public static object apply_shading;
        public static object do_not_apply;
        public static object default_table_behaviour;
        public static object auto_fit_true;
        public static OracleCommand c;
        public static char u;
        private int sheetnum;
        private bool flag_not_exported;
        public string ContractNumber { get; set; }

        static List_of_dictionaries()
        {
            List_of_dictionaries.oWord = (Word.Application)null;
            List_of_dictionaries.oDoc = (_Document)null;

            List_of_dictionaries.oEndOfDoc = "\\endofdoc";
            List_of_dictionaries.filename = "D:\\hello.rtf";
            List_of_dictionaries.m_objExcel = (Microsoft.Office.Interop.Excel.Application)null;
            List_of_dictionaries.m_objBooks = (Workbooks)null;
            List_of_dictionaries.m_objBook = (_Workbook)null;
            List_of_dictionaries.m_objSheets = (Sheets)null;
            List_of_dictionaries.m_objSheet = (_Worksheet)null;
            List_of_dictionaries.m_objRange = (Microsoft.Office.Interop.Excel.Range)null;
            List_of_dictionaries.m_objFont = (Microsoft.Office.Interop.Excel.Font)null;
            List_of_dictionaries.m_objstyle = (Microsoft.Office.Interop.Excel.Style)null;
            List_of_dictionaries.m_objOpt = Missing.Value;
            List_of_dictionaries.col_width = 200;
            List_of_dictionaries.direction = WdDocumentDirection.wdLeftToRight;
            List_of_dictionaries.format = WdTableFormat.wdTableFormatGrid3;
            List_of_dictionaries.tab_direction = WdTableDirection.wdTableDirectionLtr;
            List_of_dictionaries.auto_fit = WdAutoFitBehavior.wdAutoFitContent;
            List_of_dictionaries.auto_border = WdBorderType.wdBorderHorizontal;
            List_of_dictionaries.auto_heading = WdHeadingSeparator.wdHeadingSeparatorLetterFull;
            List_of_dictionaries.apply_color = WdTableFormatApply.wdTableFormatApplyColor;
            List_of_dictionaries.apply_header = WdTableFormatApply.wdTableFormatApplyHeadingRows;
            List_of_dictionaries.apply_shading = WdTableFormatApply.wdTableFormatApplyShading;
            List_of_dictionaries.do_not_apply = false;
            List_of_dictionaries.default_table_behaviour = WdDefaultTableBehavior.wdWord9TableBehavior;
            List_of_dictionaries.auto_fit_true = true;
            List_of_dictionaries.c = Frm_1.dbcon.CreateCommand();
        }

        public List_of_dictionaries()
        {
            this.components = (IContainer)null;
            this.contract_type = "select type Code ,Name,decode(status,'1','Active','Inactive') Status  from a4m.tcontracttype where branch=" + Frm_1.bank_num + " order by 1";
            this.contract_status = "select tcontractstatereference.STATEID ID,tcontractstatereference.STATECODE Code,tcontractstatereference.STATENAME Name from a4m.tcontractstatereference where branch=" + Frm_1.bank_num + " order by 1";
            this.Account_type = "select Accounttype Code ,Name,decode(flags,1,'Active','Inactive') Status  from a4m.taccounttype where branch=" + Frm_1.bank_num + " order by 1";
            this.card_financial_profiles = "select profile Code,Name Fee_Profile,decode(status,1,'Active','Inactive') Status from a4m.tfinprofile where branch=" + Frm_1.bank_num + " and activity=1 order by 1";
            this.Retailers_financial_profiles = "select profile Code,Name Fee_Profile,decode(status,1,'Active','Inactive') Status from a4m.tfinprofile where branch=" + Frm_1.bank_num + " and activity=2 order by 1";
            this.Telebank_financial_profiles = "select profile Code,Name Fee_Profile,decode(status,1,'Active','Inactive') Status from a4m.tfinprofile where branch=" + Frm_1.bank_num + " and activity=3 order by 1";
            this.SMS_financial_profiles = "select profile Code,Name Fee_Profile,decode(status,1,'Active','Inactive') Status from a4m.tfinprofile where branch=" + Frm_1.bank_num + " and activity=6 order by 1";
            this.SMS_Channels_financial_profiles = "select profile Code,Name Fee_Profile,decode(status,1,'Active','Inactive') Status from a4m.tfinprofile where branch=" + Frm_1.bank_num + " and activity=7 order by 1";
            this.stat_lang = "select Code ,Name from a4m.tstatementlanguage where branch=" + Frm_1.bank_num + " order by 1";
            this.limit_groups = "select ID Code,GRPNAME Name,decode(objtype,1,'Card',2,'Account',3,'Acc-To-Card',4,'Telebank') Limit_TYpe from a4m.tlimitgrp where branch=" + Frm_1.bank_num + " order by 1";
            this.usage_limits = "select LIMITID Code,Name,decode(Limittype,1,'Counter',2,'Amount','Others') Type,(decode(substr(Objtypemask,1,1) ,'1','Card ',' ')||decode(substr(Objtypemask,2,1) ,'1','Account ',' ')||decode(substr(Objtypemask,3,1) ,'1','Acct-To-Card ',' ')||decode(substr(Objtypemask,4,1) ,'1','Telebank',' ')) Limit_Link from a4m.treferencecardlimit where branch=" + Frm_1.bank_num + " order by 1";
            this.countries = "SELECT treferencecountry.Code Code, treferencecountry.ABBREVIATION2 ABBREVIATION2, treferencecountry.ABBREVIATION3 ABBREVIATION3, treferencecountry.NAME Name, treferencecountry.DESCRIPTION Description, treferencecountry.PHONE Phone, tgroupcountry.ident Related_Country_Group_ID, tgroupcountry.code Related_Country_Code, A4M.TREFERENCEGROUPCOUNTRY.ABBREV, tgroupcountry.network Related_Country_Network FROM a4m.treferencecountry, a4m.tgroupcountry, A4M.TREFERENCEGROUPCOUNTRY WHERE treferencecountry.branch = tgroupcountry.branch AND treferencecountry.code = tgroupcountry.code AND treferencecountry.branch = " + Frm_1.bank_num + " AND A4M.TREFERENCEGROUPCOUNTRY.branch = a4m.tgroupcountry.branch AND A4M.TREFERENCEGROUPCOUNTRY.network = a4m.tgroupcountry.network AND A4M.TREFERENCEGROUPCOUNTRY.ident = a4m.tgroupcountry.ident AND treferencecountry.Code = tgroupcountry.code ORDER BY 1";
            this.currencies = "select treferencecurrency.currency currency,treferencecurrency.description description,treferencecurrency.abbreviation abbreviation,treferencecurrency.precision precision,treferencecurrency.favorite favorite from a4m.treferencecurrency where branch=" + Frm_1.bank_num + " order by 1";
            this.regions = "select Code,Name from a4m.treferenceregion where branch=" + Frm_1.bank_num + " order by 1";
            this.cities = "select t1.code,t1.name,t2.code related_region_code,t2.name related_region_name from a4m.treferencecity t1,a4m.treferenceregion t2 where t1.branch=t2.branch and t1.region=t2.code and t1.branch=" + Frm_1.bank_num + " order by 1";
            this.min_payments = "select tcontractmpprofile.PROFILEID Code,tcontractmpprofile.PROFILENAME  from a4m.tcontractmpprofile where branch=" + Frm_1.bank_num + " order by 1";
            this.calculation_profiles = "select tcontractprofile.PROFILEID Code,tcontractprofile.PROFILENAME,tcontractprofile.CURRENCY from a4m.tcontractprofile where branch=" + Frm_1.bank_num + " order by 1";
            this.direct_debit_profile = "select tcontractddreference.profileid Code,tcontractddreference.profilename Fee_Profile from a4m.tcontractddreference where branch=" + Frm_1.bank_num + " order by 1";
            this.occupation = "select Code,Name from a4m.toccupation where branch=" + Frm_1.bank_num + " order by 1";
            this.branches = "select branchpart Code,Name,Ident External_Code from a4m.tbranchpart where branch=" + Frm_1.bank_num + " order by 1";
            this.card_status = "select treferencecrd_stat.crd_stat code,name from a4m.treferencecrd_stat order by 1";
            this.card_state = "select treferencecardsign.cardsign code,treferencecardsign.name from a4m.treferencecardsign where branch=" + Frm_1.bank_num + " order by 1";
            this.marital_status = "select tfamilyscore.code,tfamilyscore.state Name from a4m.tfamilyscore where branch=" + Frm_1.bank_num + " order by 1";
            this.interest_rate = "select tpercentname.id Code,tpercentname.NAME from a4m.tpercentname where branch=" + Frm_1.bank_num + " order by 1";
            this.card_product = "select a4m.treferencecardproduct.CODE, a4m.treferencecardproduct.NAME,prefix BIN,period Validity_Months,ServiceCODE,A4M.tlimitgrp.GRPNAME Usage_Limit,A4M.TFINPROFILE.NAME Financial_Profile,A4M.TREFERENCECARDSIGN.NAME State_upon_Creation,a4m.referencecrd_stat.getname(a4m.treferencecardproduct.Status) Status_upon_Creation,(select '*'  from a4m.tgroupcp where branch=" + Frm_1.bank_num + " and code=treferencecardproduct.code and upper(ident) ='MC PRODUCTS' ) as MC_Product,(select '*' from a4m.tgroupcp where branch=" + Frm_1.bank_num + " and code=treferencecardproduct.code and upper(ident) ='VISA PRODUCTS' )  as VISA_Product,(select '*'  from a4m.tgroupcp where branch=" + Frm_1.bank_num + " and code=treferencecardproduct.code and upper(ident) ='DEBIT PRODUCTS' ) as Debit_Product,(select '*'  from a4m.tgroupcp where branch=" + Frm_1.bank_num + " and code=treferencecardproduct.code and upper(ident) ='PREPAID PRODUCTS' ) as Prepaid_Product,(select '*'  from a4m.tgroupcp where branch=" + Frm_1.bank_num + " and code=treferencecardproduct.code and upper(ident) ='CREDIT PRODUCTS' ) as Credit_Product,(select '*' from a4m.tgroupcp where branch=" + Frm_1.bank_num + " and code=treferencecardproduct.code and upper(ident) ='CORPORATE PRODUCTS' ) as Corporate_Product ,(select '*' from a4m.tgroupcp where branch=" + Frm_1.bank_num + " and code=treferencecardproduct.code and upper(ident) ='VIRTUAL PRODUCTS' )  as Virtual_Product from a4m.treferencecardproduct,a4m.tlimitgrp,a4m.tfinprofile,a4m.treferencecardsign where a4m.treferencecardproduct.branch=" + Frm_1.bank_num + " and a4m.tfinprofile.branch(+)=a4m.treferencecardproduct.branch and a4m.tfinprofile.profile(+)=a4m.treferencecardproduct.finprofile and a4m.treferencecardsign.branch(+)=a4m.treferencecardproduct.branch and a4m.treferencecardsign.cardsign(+)=a4m.treferencecardproduct.state and a4m.tlimitgrp.branch(+)=a4m.treferencecardproduct.branch and a4m.tlimitgrp.id(+)=a4m.treferencecardproduct.GRPLimitID order by 1";
            this.titles = "select titlecode code,titletext name from a4m.tclientpersonetitle where branch=" + Frm_1.bank_num + " order by 1";
            this.pan_ranges = "SELECT rp.NAME,rp.prefix,pr.onum Field_NO,DECODE (pr.ftype,1, 'Sequence',3, 'Customer ID',2, 'Card Serial No.',5, 'Constant',6, 'Random Number',7, 'Card Sub-Branch Code (BranchPart)',8, 'Customer External ID',9, 'Customer Sub-Branch Code (BranchPart)',10, 'Sequence With Regard To Branches',11, 'Random Number With Regard To Branches',12, 'Sequence (Ranges)',13, 'Random Number (Ranges)',14, 'PL/SQL Block',NULL) TYPE,pran.no,nvl(pran.SEQMIN,pr.SEQMIN) MIN,nvl(pran.SEQMAX,pr.SEQMAX) MAX,decode(Active,1,'*',' ') Active,pran.minparent Branch_From,pran.maxparent Branch_To,(case(nvl(pran.maxparent,0))when 0 then round(decode(pran.CURSEQ,0,0,((((pran.CURSEQ-pran.SEQMIN)+1)*100)/((pran.SEQMAX-pran.SEQMIN)+1))),1) else null end) Current_Usage_Percentage FROM a4m.treferencecardproduct rp, a4m.tcpcardnumber pr,a4m.tcpcardnumberranges pran WHERE rp.branch = pr.branch AND rp.code = pr.code AND rp.branch =" + Frm_1.bank_num + " and pr.FIELDCODE=pran.FIELDCODE(+) ORDER BY 2,3,5";
            this.bank_operators = "select t1.name UserID,t3.fio UserName,t5.NAME PrivilegeGroup,decode(t1.type,0,'Active',1,'Deleted') UserStatus from a4m.tclerk t1,a4m.tseance t2,a4m.tclientpersone t3,a4m.tclerk2group t4,a4m.tclerkgroup t5 where t1.branch= " + Frm_1.bank_num + " and t1.branch=t2.branch and (t1.NAME like t2.FIID||'%' or t1.name like 'BMISR%') and t1.branch=t3.branch and t1.CLERKID=t3.IDCLIENT and t1.branch=t4.branch(+) and t1.CODE=t4.CLERKID(+) and t4.branch=t5.branch(+) and t4.GROUPID=t5.ID(+) order by UserID ";
            this.mscc_operators = "select t1.name UserID,t3.fio UserName,t5.NAME PrivilegeGroup,decode(t1.type,0,'Active',1,'Deleted') UserStatus from a4m.tclerk t1,a4m.tseance t2,a4m.tclientpersone t3,a4m.tclerk2group t4,a4m.tclerkgroup t5 where t1.branch= " + Frm_1.bank_num + " and t1.branch=t2.branch and (t1.NAME not like t2.FIID||'%' or t1.name not like 'BMISR%') and t1.branch=t3.branch and t1.CLERKID=t3.IDCLIENT and t1.branch=t4.branch(+) and t1.CODE=t4.CLERKID(+) and t4.branch=t5.branch(+) and t4.GROUPID=t5.ID(+) order by UserID ";
            this.Dict_Schema = "select t1.Code,t1.ID,t1.Description from a4m.tcnsrefschema t1 where t1.branch= " + Frm_1.bank_num + "  order by 1 ";
            this.Dict_Channel = "select t1.Code,t1.ID,t1.Description from a4m.tcnsrefchannel t1 where t1.branch= " + Frm_1.bank_num + "  order by 1 ";
            this.Education = "select t1.Code,t1.Education from a4m.teducationscore t1 where t1.branch= " + Frm_1.bank_num + "  order by 1 ";
            this.country_groups = "select t1.ident ID,t1.Name Name from a4m.treferencegroupcountry t1 where branch=" + Frm_1.bank_num + " order by 1";
            this.mcc_groups = "select t1.ident ID,t1.name from a4m.treferencegroupmcc t1 where branch=" + Frm_1.bank_num + " order by 1";
            //Mamr - old 19-7-2020
            //this.mcc = "select t.mcc,t.name,decode(t.favorite,'1','YES',null) as Is_Favorite ,(select name from a4m.tgroupmcc,a4m.treferencegroupmcc where tgroupmcc.branch= treferencegroupmcc.branch and tgroupmcc.ident= treferencegroupmcc.ident and tgroupmcc.branch=" + Frm_1.bank_num + " and mcc=t.mcc and upper(tgroupmcc.ident) like 'V%') as VISA_MCC_Group,(select name from a4m.tgroupmcc,a4m.treferencegroupmcc where tgroupmcc.branch= treferencegroupmcc.branch and tgroupmcc.ident= treferencegroupmcc.ident and tgroupmcc.branch=" + Frm_1.bank_num + " and mcc=t.mcc and upper(tgroupmcc.ident) like 'E%') as MC_MCC_Group from a4m.treferencemcc t where branch=" + Frm_1.bank_num + " order by 1";
            //Mamr - new 19-7-2020 resolve subquery more than one row issue
            this.mcc = "select t.mcc,t.name,decode(t.favorite,'1','YES',null) as Is_Favorite ,(select name from a4m.tgroupmcc,a4m.treferencegroupmcc where tgroupmcc.branch= treferencegroupmcc.branch and tgroupmcc.ident= treferencegroupmcc.ident and tgroupmcc.branch=" + Frm_1.bank_num + " and mcc=t.mcc and upper(tgroupmcc.ident) like 'V%' and rownum <= 1) as VISA_MCC_Group,(select name from a4m.tgroupmcc,a4m.treferencegroupmcc where tgroupmcc.branch= treferencegroupmcc.branch and tgroupmcc.ident= treferencegroupmcc.ident and tgroupmcc.branch=" + Frm_1.bank_num + " and mcc=t.mcc and upper(tgroupmcc.ident) like 'E%' and rownum <= 1) as MC_MCC_Group from a4m.treferencemcc t where branch=" + Frm_1.bank_num + " order by 1";
            this.card_USA = "select EXTID CODE,DESCRIPTION NAME,DECODE(UPPER(PTYPE),'STR','STRING','NUM','NUMBER',UPPER(PTYPE)) TYPE,LEN LENGTH,DECODE(nvl(MANDAT,0),0,'NO',1,'YES') IS_MANDATORY,DECODE(nvl(EDITABLE,0),0,'NO',1,'YES') IS_EDITABLE,DECODE(nvl(ISDEFAULT,0),0,'NO',1,'YES',2,'YES') HAS_DEFAULT,DECODE(UPPER(PTYPE),'STR',VALUEDEFSTR,'NUM',VALUEDEFNUM,'DATE',VALUEDEFDATE) DEFAULT_VALUE,DECODE(nvl(ISLISTVAL,0),0,'NO',1,'YES') HAS_LIST FROM A4M.TOBJADDITIONALPROPERTY where BRANCH=" + Frm_1.bank_num + " and OBJECTID='CARD' order by SORT_NO";
            this.account_USA = "select EXTID CODE,DESCRIPTION NAME,DECODE(UPPER(PTYPE),'STR','STRING','NUM','NUMBER',UPPER(PTYPE)) TYPE,LEN LENGTH,DECODE(nvl(MANDAT,0),0,'NO',1,'YES') IS_MANDATORY,DECODE(nvl(EDITABLE,0),0,'NO',1,'YES') IS_EDITABLE,DECODE(nvl(ISDEFAULT,0),0,'NO',1,'YES',2,'YES') HAS_DEFAULT,DECODE(UPPER(PTYPE),'STR',VALUEDEFSTR,'NUM',VALUEDEFNUM,'DATE',VALUEDEFDATE) DEFAULT_VALUE,DECODE(nvl(ISLISTVAL,0),0,'NO',1,'YES') HAS_LIST FROM A4M.TOBJADDITIONALPROPERTY where BRANCH=" + Frm_1.bank_num + " and OBJECTID='ACCOUNT' order by SORT_NO";
            this.contract_USA = "select EXTID CODE,DESCRIPTION NAME,DECODE(UPPER(PTYPE),'STR','STRING','NUM','NUMBER',UPPER(PTYPE)) TYPE,LEN LENGTH,DECODE(nvl(MANDAT,0),0,'NO',1,'YES') IS_MANDATORY,DECODE(nvl(EDITABLE,0),0,'NO',1,'YES') IS_EDITABLE,DECODE(nvl(ISDEFAULT,0),0,'NO',1,'YES',2,'YES') HAS_DEFAULT,DECODE(UPPER(PTYPE),'STR',VALUEDEFSTR,'NUM',VALUEDEFNUM,'DATE',VALUEDEFDATE) DEFAULT_VALUE,DECODE(nvl(ISLISTVAL,0),0,'NO',1,'YES') HAS_LIST FROM A4M.TOBJADDITIONALPROPERTY where BRANCH=" + Frm_1.bank_num + " and OBJECTID='CONTRACT' order by SORT_NO";
            this.client_USA = "select EXTID CODE,DESCRIPTION NAME,DECODE(UPPER(PTYPE),'STR','STRING','NUM','NUMBER',UPPER(PTYPE)) TYPE,LEN LENGTH,DECODE(nvl(MANDAT,0),0,'NO',1,'YES') IS_MANDATORY,DECODE(nvl(EDITABLE,0),0,'NO',1,'YES') IS_EDITABLE,DECODE(nvl(ISDEFAULT,0),0,'NO',1,'YES',2,'YES') HAS_DEFAULT,DECODE(UPPER(PTYPE),'STR',VALUEDEFSTR,'NUM',VALUEDEFNUM,'DATE',VALUEDEFDATE) DEFAULT_VALUE,DECODE(nvl(ISLISTVAL,0),0,'NO',1,'YES') HAS_LIST FROM A4M.TOBJADDITIONALPROPERTY where BRANCH=" + Frm_1.bank_num + " and OBJECTID='CLIENT' order by SORT_NO";
            this.sms_dic = "select code,name,active from a4m.tStmtNotifyTemplate where branch =" + Frm_1.bank_num;
            this.BaseInstallment = $"select D.NAME,D.TYPE from A4M.TCONTRACTTYPE d where branch = {Frm_1.bank_num} and SCHEMATYPE = 3";
            this.BaseInstallmentSettings = $"SELECT KEY, DECODE(KEY, 'DAYSINYEAR', DECODE(VALUE, '1', 'Native','2', '360/30', VALUE), 'COUNTDAYSMODE', DECODE(VALUE, '1', 'days in cycle','2', 'days in month', VALUE), 'CALCMETHOD', DECODE(VALUE, '1', 'Initial balance, fixed repayment','2', 'Unpaid balance, fixed repayment','3','Unpaid balance, decreasing repayment','4','Unpaid balance, fixed repayment per month', VALUE), 'CALCSCHEDULEBY', DECODE(VALUE, '1', 'Count of billing cycles','2', 'Amount of regular repayment', VALUE), 'INSTALLMENTCT', DECODE(VALUE, '1', 'Charge interest','2', 'As fee', '3', 'As prorated fee',VALUE), 'BASE2ACCELERATION', DECODE(VALUE, '1', 'Dont accrue','2', 'Accrue', VALUE), 'PENALTYCHARGETYPE', DECODE(VALUE, '1', 'Do not charge','2', 'Charge as interest','3','harge as fee', VALUE), 'LIMITMETHOD', DECODE(VALUE, '1', 'As total of','2', 'As maximum between','3','As minimum between', VALUE), 'ACCELDATEID', DECODE(VALUE, '1', 'Current Date','2', 'Statement Date', VALUE), 'MULTIPLELOAN', DECODE(VALUE, '1', 'Allowed','2', 'Not allowed', VALUE), 'FEEMODE', DECODE(VALUE, '1', 'Do not charge','2', 'From RCM settings','3','From Installment settings','4','RCM and Installment', VALUE), 'FEEAMOUNT', DECODE(VALUE, '1', 'Remaining debt','2', 'Full amount','3','Unpaid principal amount', VALUE), 'FIXEDINTEREST', DECODE(VALUE, '1', 'New tranches only','2', 'All tranches', VALUE), VALUE) AS VALUE, D.CONTRACTTYPE FROM A4m.TCONTRACTTYPEPARAMETERS d where branch = {Frm_1.bank_num} and and CONTRACTTYPE = {ContractNumber}";
            this.LinkedContractTypes = $"select CONTRACT.NAME from A4M.TCONTRACTTYPE contract join A4M.TCONTRACTTYPELink link on LINK.MAINTYPE = CONTRACT.TYPE where CONTRACT.BRANCH = {Frm_1.bank_num} and SCHEMATYPE != 3 and LINKTYPE = {ContractNumber}";
            this.sheetnum = 1;
            this.flag_not_exported = false;
            //base.\u002Ector();


            this.InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && this.components != null)
                this.components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.button2 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.checkedListBox2 = new System.Windows.Forms.CheckedListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.BackColor = System.Drawing.SystemColors.Window;
            this.checkedListBox1.CheckOnClick = true;
            this.checkedListBox1.ColumnWidth = 100;
            this.checkedListBox1.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkedListBox1.ForeColor = System.Drawing.Color.MidnightBlue;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.HorizontalScrollbar = true;
            this.checkedListBox1.Items.AddRange(new object[] {
            "Dictionary of Contract Types",
            "Dictionary of Contract Status",
            "Dictionary of Account Types",
            "Dictionary of Card Financial Profiles",
            "Dictionary of Retailers Financial Profiles",
            "Dictionary of Telebank Financial Profiles",
            "Dictionary of SMS Financial Profiles",
            "Dictionary of SMS Channels Financial Profiles",
            "Dictionary of SMS Notification Templates",
            "Dictionary of Limit Groups",
            "Dictionary of Usage Limits",
            "Dictionary of Countries",
            "Dictionary of Regions",
            "Dictionary of Cities",
            "Dictionary of Currencies",
            "Dictionary of Minimum Payment",
            "Dictionary of Calculation Profile",
            "Dictionary of Direct Debit Profile",
            "Dictionary of Interest Rate",
            "Dictionary of Online Status",
            "Dictionary of CMS States",
            "Dictionary of Card Products",
            "Dictionary of Pan Ranges",
            "Dictionary of Bank Operators",
            "Dictionary of MSCC Operators",
            "Dictionary of Marital Status",
            "Dictionary of Titles",
            "Dictionary of Occupation",
            "Dictionary of Education",
            "Dictionary of Branches",
            "Dictionary of Card user attributes",
            "Dictionary of Account user attributes",
            "Dictionary of Contract user attributes",
            "Dictionary of Client user attributes",
            "Dictionary of MCC",
            "Dictionary of CMS Schemes",
            "Dictionary of CMS Channels"});
            this.checkedListBox1.Location = new System.Drawing.Point(2, 25);
            this.checkedListBox1.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(310, 346);
            this.checkedListBox1.TabIndex = 0;
            this.checkedListBox1.SelectedIndexChanged += new System.EventHandler(this.checkedListBox1_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Silver;
            this.label1.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.label1.Location = new System.Drawing.Point(-1, 0);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(199, 14);
            this.label1.TabIndex = 1;
            this.label1.Text = "Available List Of Dictionaries";
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.RestoreDirectory = true;
            // 
            // linkLabel1
            // 
            this.linkLabel1.ActiveLinkColor = System.Drawing.Color.Sienna;
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.BackColor = System.Drawing.Color.Silver;
            this.linkLabel1.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabel1.ForeColor = System.Drawing.Color.DarkGreen;
            this.linkLabel1.Location = new System.Drawing.Point(-1, 378);
            this.linkLabel1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(60, 13);
            this.linkLabel1.TabIndex = 13;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "Select All";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.Linen;
            this.button2.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.button2.Image = global::Issues_Application.Properties.Resources.form_compile;
            this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button2.Location = new System.Drawing.Point(493, 383);
            this.button2.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(227, 27);
            this.button2.TabIndex = 14;
            this.button2.Text = "Generate Settings Report";
            this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button4
            // 
            this.button4.BackColor = System.Drawing.Color.Linen;
            this.button4.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button4.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.button4.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button4.Location = new System.Drawing.Point(698, 448);
            this.button4.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(50, 22);
            this.button4.TabIndex = 11;
            this.button4.Text = "Exit";
            this.button4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button4.UseVisualStyleBackColor = false;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Linen;
            this.button1.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button1.Location = new System.Drawing.Point(80, 383);
            this.button1.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(227, 27);
            this.button1.TabIndex = 2;
            this.button1.Text = "Generate Dictionaries File";
            this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.Linen;
            this.button3.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.button3.Image = global::Issues_Application.Properties.Resources.form_compile;
            this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button3.Location = new System.Drawing.Point(290, 419);
            this.button3.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(228, 27);
            this.button3.TabIndex = 15;
            this.button3.Text = "Generate Service Code Form";
            this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Visible = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // checkedListBox2
            // 
            this.checkedListBox2.BackColor = System.Drawing.SystemColors.Window;
            this.checkedListBox2.CheckOnClick = true;
            this.checkedListBox2.ColumnWidth = 100;
            this.checkedListBox2.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkedListBox2.ForeColor = System.Drawing.Color.MidnightBlue;
            this.checkedListBox2.FormattingEnabled = true;
            this.checkedListBox2.HorizontalScrollbar = true;
            this.checkedListBox2.Items.AddRange(new object[] {
            "Section 1 : Financial Institution Details",
            "Section 2 : Card Products",
            "Section 3 : Card Profiles",
            "Section 4 : Periodic Usage Limit",
            "Section 5 : Delinquency Settings (Credit Products Only)",
            "Section 6 : Working Calender and Billing Cycle Calender",
            "Section 7 : Interest Settings (Credit Products Only)",
            "Section 8 : Allowable Overlimit + Overlimit Fees (Credit Products Only)",
            "Section 9 : Allowable Overdue + Overdue Fees (Credit Products Only)",
            "Section 10 : Credit Shield (Credit Products Only)",
            "Section 11 : Minimum Payment and Direct Debit Settings(Credit Products Only)",
            "Section 12 : Card Limits (Credit Products Only)",
            "Section 13 : Installment Setting",
            "Section 14 : Appendix"});
            this.checkedListBox2.Location = new System.Drawing.Point(328, 25);
            this.checkedListBox2.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.checkedListBox2.Name = "checkedListBox2";
            this.checkedListBox2.Size = new System.Drawing.Size(393, 346);
            this.checkedListBox2.TabIndex = 16;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Silver;
            this.label2.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.label2.Location = new System.Drawing.Point(334, 0);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(169, 14);
            this.label2.TabIndex = 17;
            this.label2.Text = "List Of Settings Sections";
            // 
            // linkLabel2
            // 
            this.linkLabel2.ActiveLinkColor = System.Drawing.Color.Sienna;
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.BackColor = System.Drawing.Color.Silver;
            this.linkLabel2.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabel2.ForeColor = System.Drawing.Color.DarkGreen;
            this.linkLabel2.Location = new System.Drawing.Point(330, 378);
            this.linkLabel2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(60, 13);
            this.linkLabel2.TabIndex = 18;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Text = "Select All";
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
            this.checkBox1.ForeColor = System.Drawing.Color.MediumOrchid;
            this.checkBox1.Location = new System.Drawing.Point(326, 393);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(167, 17);
            this.checkBox1.TabIndex = 19;
            this.checkBox1.Text = "Include Settlement Flags";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // List_of_dictionaries
            // 
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(748, 469);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.linkLabel2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.checkedListBox2);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.checkedListBox1);
            this.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.MaximizeBox = false;
            this.Name = "List_of_dictionaries";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "List Of Dictionaries";
            this.Load += new System.EventHandler(this.List_of_dictionaries_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void List_of_dictionaries_Load(object sender, EventArgs e)
        {
            this.label1.Text = "List Of Dictionaries For " + Frm_1.bank_fiid + " On " + (Frm_1.dbcon.DataSource).ToString();
            this.sheetnum = 1;
            AssemblyInfo inf = new AssemblyInfo();
            this.Text = this.Text + " " + System.Windows.Forms.Application.ProductName + " - Version:" + inf.Version;
        }


        private void execute_proc(string command, string sheetname, Color col, bool character)
        {
            try
            {
                List_of_dictionaries.c.CommandText = command;
                OracleDataReader oracleDataReader = List_of_dictionaries.c.ExecuteReader();
                if (this.sheetnum > 1)
                {
                    List_of_dictionaries.m_objSheet = (_Worksheet)List_of_dictionaries.m_objSheets.Add(List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt);
                    List_of_dictionaries.m_objSheet.Name = sheetname;
                    List_of_dictionaries.m_objSheet.Tab.Color = ColorTranslator.ToOle(col);
                    ++this.sheetnum;
                }
                else
                {
                    List_of_dictionaries.m_objSheet = (_Worksheet)List_of_dictionaries.m_objSheets.get_Item(this.sheetnum);
                    List_of_dictionaries.m_objSheet.Name = sheetname;
                    List_of_dictionaries.m_objSheet.Tab.Color = ColorTranslator.ToOle(col);
                    ++this.sheetnum;
                }
                for (int ordinal = 0; ordinal < oracleDataReader.FieldCount; ++ordinal)
                {
                    char ch = (char)((uint)'A' + (uint)ordinal);
                    List_of_dictionaries.m_objRange = List_of_dictionaries.m_objSheet.get_Range((ch.ToString() + "1"), List_of_dictionaries.m_objOpt);
                    List_of_dictionaries.m_objRange.set_Value(null, (oracleDataReader.GetName(ordinal)).ToString());
                    List_of_dictionaries.u = ch;
                }
                int num1 = 1;
                List_of_dictionaries.m_objRange = List_of_dictionaries.m_objSheet.get_Range("A1", (List_of_dictionaries.u.ToString() + "1"));
                List_of_dictionaries.m_objstyle = List_of_dictionaries.m_objBook.Styles.Add("NewStyle" + this.sheetnum.ToString(), List_of_dictionaries.m_objOpt);
                List_of_dictionaries.m_objstyle.Font.Name = "Verdana";
                List_of_dictionaries.m_objstyle.Font.Size = 10;
                List_of_dictionaries.m_objstyle.Font.Color = ColorTranslator.ToOle(Color.White);
                List_of_dictionaries.m_objstyle.Interior.Color = ColorTranslator.ToOle(Color.Blue);
                List_of_dictionaries.m_objstyle.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid;
                List_of_dictionaries.m_objstyle.Font.Bold = true;
                List_of_dictionaries.m_objstyle.Font.Underline = true;
                List_of_dictionaries.m_objRange.Justify();
                List_of_dictionaries.m_objRange.ColumnWidth = 30;
                List_of_dictionaries.m_objRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                List_of_dictionaries.m_objRange.Style = ("NewStyle" + this.sheetnum.ToString());
                while (oracleDataReader.Read())
                {
                    ++num1;
                    for (int index = 0; index < oracleDataReader.FieldCount; ++index)
                    {
                        char ch = (char)((uint)'A' + (uint)index);
                        List_of_dictionaries.m_objRange = List_of_dictionaries.m_objSheet.get_Range((ch.ToString() + num1.ToString()), List_of_dictionaries.m_objOpt);
                        if (character)
                            List_of_dictionaries.m_objRange.set_Value(null, ("'" + oracleDataReader[index].ToString()));
                        else
                            List_of_dictionaries.m_objRange.set_Value(null, oracleDataReader[index].ToString());
                        List_of_dictionaries.m_objRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    }
                }
                oracleDataReader.Close();
                if (num1 != 1)
                    return;
                int num2 = num1 + 1;
                List_of_dictionaries.m_objRange = List_of_dictionaries.m_objSheet.get_Range(("A" + num2.ToString()), List_of_dictionaries.m_objOpt);
                List_of_dictionaries.m_objRange.set_Value(null, "Not Defined");
                List_of_dictionaries.m_objRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            }
            catch (System.Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                this.flag_not_exported = true;
            }
        }

        private void generate(string filename)
        {
            List_of_dictionaries.m_objExcel = (Microsoft.Office.Interop.Excel.Application)new Microsoft.Office.Interop.Excel.Application();
            List_of_dictionaries.m_objBooks = List_of_dictionaries.m_objExcel.Workbooks;
            List_of_dictionaries.m_objBook = (_Workbook)List_of_dictionaries.m_objBooks.Add(List_of_dictionaries.m_objOpt);
            List_of_dictionaries.m_objSheets = List_of_dictionaries.m_objBook.Worksheets;
            for (int index = this.checkedListBox1.Items.Count - 1; index >= 0; --index)
            {
                if (this.checkedListBox1.GetItemChecked(index))
                {
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Contract Types")
                        this.execute_proc(this.contract_type, "Dict. of Contract Types", Color.Snow, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Contract Status")
                        this.execute_proc(this.contract_status, "Dict. of Contract Status", Color.SpringGreen, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Account Types")
                        this.execute_proc(this.Account_type, "Dict. of Account Types", Color.SkyBlue, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Card Financial Profiles")
                        this.execute_proc(this.card_financial_profiles, "Dict. of Card Fin. Prof.", Color.Lavender, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Retailers Financial Profiles")
                        this.execute_proc(this.Retailers_financial_profiles, "Dict. of Retailers Fin. Prof.", Color.Lavender, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Telebank Financial Profiles")
                        this.execute_proc(this.Telebank_financial_profiles, "Dict. of Telebank Fin. Prof.", Color.Lavender, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of SMS Financial Profiles")
                        this.execute_proc(this.SMS_financial_profiles, "Dict. of SMS Fin. Prof.", Color.Lavender, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of SMS Channels Financial Profiles")
                        this.execute_proc(this.SMS_Channels_financial_profiles, "Dict. of SMS Channels Fin. Pro.", Color.Lavender, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of SMS Notification Templates")
                        this.execute_proc(this.sms_dic, "Dict. of SMS Not Temp", Color.AntiqueWhite, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Statement Language")
                        this.execute_proc(this.stat_lang, "Dict. of Statement Language", Color.Violet, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Limit Groups")
                        this.execute_proc(this.limit_groups, "Dict. of Limit Groups", Color.YellowGreen, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Usage Limits")
                        this.execute_proc(this.usage_limits, "Dict. of Usage Limits", Color.Tan, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Countries")
                        this.execute_proc(this.countries, "Dict. of Countries", Color.Crimson, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Country Groups")
                        this.execute_proc(this.country_groups, "Dict. of Country Groups", Color.Crimson, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Regions")
                        this.execute_proc(this.regions, "Dict. of Regions", Color.LightGreen, true);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Cities")
                        this.execute_proc(this.cities, "Dict. of Cities", Color.ForestGreen, true);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Currencies")
                        this.execute_proc(this.currencies, "Dict. of Currencies", Color.AliceBlue, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Minimum Payment")
                        this.execute_proc(this.min_payments, "Dict. of Minimum Payment", Color.Salmon, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Calculation Profile")
                        this.execute_proc(this.calculation_profiles, "Dict. of Calculation Profile", Color.OrangeRed, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Direct Debit Profile")
                        this.execute_proc(this.direct_debit_profile, "Dict. of Direct Debit Profile", Color.Navy, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Interest Rate")
                        this.execute_proc(this.interest_rate, "Dict. of Interest Rate", Color.Aqua, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Online Status")
                        this.execute_proc(this.card_status, "Status in Online System", Color.Gold, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of CMS States")
                        this.execute_proc(this.card_state, "States in Card Management Sys.", Color.Beige, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Occupation")
                        this.execute_proc(this.occupation, "Dict. of Occupation", Color.MediumOrchid, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Branches")
                        this.execute_proc(this.branches, "Dict. of Branches", Color.Gray, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Marital Status")
                        this.execute_proc(this.marital_status, "Dict. of Marital Status", Color.CornflowerBlue, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Titles")
                        this.execute_proc(this.titles, "Dictionary of Titles", Color.BlueViolet, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Education")
                        this.execute_proc(this.Education, "Dictionary of Education", Color.AntiqueWhite, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Card Products")
                        this.execute_proc(this.card_product, "Dict. of Card Products", Color.AliceBlue, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Pan Ranges")
                        this.execute_proc(this.pan_ranges, "Dict. of Pan Ranges", Color.Blue, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Bank Operators")
                        this.execute_proc(this.bank_operators, "Dict. of Bank Oper.", Color.Chartreuse, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of MSCC Operators")
                        this.execute_proc(this.mscc_operators, "Dict. of MSCC Oper.", Color.BlueViolet, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Card user attributes")
                        this.execute_proc(this.card_USA, "Dict. of Card Attr.", Color.YellowGreen, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Account user attributes")
                        this.execute_proc(this.account_USA, "Dict. of Account Attr.", Color.YellowGreen, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Contract user attributes")
                        this.execute_proc(this.contract_USA, "Dict. of Contract Attr.", Color.YellowGreen, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of Client user attributes")
                        this.execute_proc(this.client_USA, "Dict. of Client Attr.", Color.YellowGreen, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of MCC")
                    {
                        try
                        {
                            this.execute_proc(this.mcc, "Dict. of MCC", Color.AntiqueWhite, false);
                        }
                        catch (System.Exception)
                        {
                            int num = (int)MessageBox.Show("MCC dic. can not be extracted (1 to 1 relation not found)", "Note", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                        }
                    }
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of CMS Schemes")
                        this.execute_proc(this.Dict_Schema, "Dict. of CMS Schemes", Color.Crimson, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of CMS Channels")
                        this.execute_proc(this.Dict_Channel, "Dict. of CMS Channels", Color.Crimson, false);
                    if (this.checkedListBox1.Items[index].ToString() == "Dictionary of MCC Groups")
                        this.execute_proc(this.mcc_groups, "Dict. of MCC Groups", Color.AntiqueWhite, false);

                }
            }
            try
            {
                List_of_dictionaries.m_objBook.Protect("Crystal_2014", List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt);
                (List_of_dictionaries.m_objBook.ActiveSheet).Protect("Crystal_2014", true);
                //List_of_dictionaries.m_objBook.SaveAs(filename, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt, XlSaveAsAccessMode.xlNoChange, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt);
                List_of_dictionaries.m_objBook.SaveAs(filename, XlFileFormat.xlOpenXMLWorkbook, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt, XlSaveAsAccessMode.xlNoChange, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt);
                List_of_dictionaries.m_objBook.Close(false, List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt);
                List_of_dictionaries.m_objExcel.Quit();
                GC.Collect();
                if (!this.flag_not_exported)
                {
                    int num1 = (int)MessageBox.Show("All List of Dictionaries Successfully Exported", "Message", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else
                {
                    int num2 = (int)MessageBox.Show("List of Dictionaries Not Completely Exported", "Message", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
            catch (System.Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message + "\nPlease Check If The Selected Execl Sheet Is Used By Another Application ...", "Message", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (!(this.saveFileDialog1.ShowDialog()).ToString().ToLower().Equals("ok"))
                    return;
                this.sheetnum = 1;
                this.flag_not_exported = false;
                this.generate(this.saveFileDialog1.FileName);
            }
            catch
            {
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.linkLabel1.Text == "Select All")
            {
                for (int index = 0; index < this.checkedListBox1.Items.Count; ++index)
                    if (index != 22) //iatta
                    {
                        this.checkedListBox1.SetItemChecked(index, true);
                        this.linkLabel1.Text = "DeSelect All";
                    }

            }
            else
            {
                if (!(this.linkLabel1.Text == "DeSelect All"))
                    return;
                for (int index = 0; index < this.checkedListBox1.Items.Count; ++index)
                    this.checkedListBox1.SetItemChecked(index, false);
                this.linkLabel1.Text = "Select All";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (!((this.saveFileDialog1.ShowDialog()).ToString().ToLower() == "ok"))
                    return;
                filename = this.saveFileDialog1.FileName;
                oWord = (Word.Application)new Word.Application();
                oWord.Visible = false;
                oWord.KeyboardLatin();
                oWord.Keyboard(2057);
                oDoc = (_Document)List_of_dictionaries.oWord.Documents.Add(ref m_objOpt, ref List_of_dictionaries.m_objOpt, ref m_objOpt, ref m_objOpt);
                oDoc.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                oDoc.PageSetup.RightMargin = 20f;
                oDoc.PageSetup.LeftMargin = 20f;
                oDoc.PageSetup.TopMargin = 20f;
                oDoc.PageSetup.BottomMargin = 20f;
                oDoc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

                oDoc.ShowGrammaticalErrors = false;
                oDoc.ShowRevisions = false;
                oDoc.ShowSpellingErrors = false;

                Paragraph oPara1 = oDoc.Content.Paragraphs.Add(ref m_objOpt);
                oPara1.Range.Text = Frm_1.bank_fiid + " Settings Report On " + Frm_1.dbname;
                oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                oPara1.Range.Font.Name = "Verdana";
                oPara1.Range.Font.Color = WdColor.wdColorDarkBlue;
                oPara1.Range.Font.Size = 15f;
                oPara1.Format.SpaceAfter = 18f;
                oPara1.Range.InsertParagraphAfter();
                oPara1.Range.Font.Bold = 0;
                oPara1.Range.Font.Size = 9f;
                oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                bool flag1 = false;
                bool flag2 = false;
                bool flag3 = false;
                bool flag4 = false;
                bool flag5 = false;
                bool flag6 = false;
                bool flag7 = false;
                bool flag8 = false;
                bool flag9 = false;
                bool flag10 = false;
                bool flag11 = false;
                bool flag12 = false;
                bool flag13 = false;
                bool flag14 = false;
                for (int index = this.checkedListBox2.Items.Count - 1; index >= 0; --index)
                {
                    if (this.checkedListBox2.GetItemChecked(index))
                    {
                        if (this.checkedListBox2.Items[index].ToString() == "Section 1 : Financial Institution Details")
                            flag1 = true;
                        else if (this.checkedListBox2.Items[index].ToString() == "Section 2 : Card Products")
                            flag2 = true;
                        else if (this.checkedListBox2.Items[index].ToString() == "Section 3 : Card Profiles")
                            flag3 = true;
                        else if (this.checkedListBox2.Items[index].ToString() == "Section 4 : Periodic Usage Limit")
                            flag4 = true;
                        else if (this.checkedListBox2.Items[index].ToString() == "Section 5 : Delinquency Settings (Credit Products Only)")
                            flag5 = true;
                        else if (this.checkedListBox2.Items[index].ToString() == "Section 6 : Working Calender and Billing Cycle Calender")
                            flag6 = true;
                        else if (this.checkedListBox2.Items[index].ToString() == "Section 7 : Interest Settings (Credit Products Only)")
                            flag7 = true;
                        else if (this.checkedListBox2.Items[index].ToString() == "Section 8 : Allowable Overlimit + Overlimit Fees (Credit Products Only)")
                            flag8 = true;
                        else if (this.checkedListBox2.Items[index].ToString() == "Section 9 : Allowable Overdue + Overdue Fees (Credit Products Only)")
                            flag9 = true;
                        else if (this.checkedListBox2.Items[index].ToString() == "Section 10 : Credit Shield (Credit Products Only)")
                            flag10 = true;
                        else if (this.checkedListBox2.Items[index].ToString() == "Section 11 : Minimum Payment and Direct Debit Settings(Credit Products Only)")
                            flag11 = true;
                        else if (this.checkedListBox2.Items[index].ToString() == "Section 12 : Card Limits (Credit Products Only)")
                            flag12 = true;
                        else if (this.checkedListBox2.Items[index].ToString() == "Section 13 : Installment Setting")
                            flag13 = true;
                        else if (this.checkedListBox2.Items[index].ToString() == "Section 14 : Appendix")
                            flag14 = true;
                    }
                }
                if (flag1)
                {
                    oPara1.Range.Text = "Section 1 : Financial Institution Details";
                    oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara1.Range.Font.Bold = 0;
                    oPara1.Range.Font.Size = 9f;
                    oPara1.Format.SpaceAfter = 0.0f;
                    oPara1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara1.Range.InsertParagraphAfter();
                    Paragraph oPara2 = oDoc.Content.Paragraphs.Add(ref m_objOpt);
                    oPara2.Range.LanguageID = WdLanguageID.wdEnglishUK;
                    oPara2.Range.Text = "\nS1:1-Bank Basic Details\n";
                    oPara2.Range.LanguageID = WdLanguageID.wdEnglishUK;
                    oPara2.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara2.Range.InsertParagraphAfter();
                    oPara2.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range1 = oDoc.Content.Paragraphs.Add(ref m_objOpt).Range;
                    range1.Text = "Bank Name\tBank Fiid\n";
                    Microsoft.Office.Interop.Word.Range range2 = range1;
                    string str1 = range2.Text + Frm_1.bank_name + "\t" + Frm_1.bank_fiid + "\n";
                    range2.Text = str1;
                    range1.DetectLanguage();
                    range1.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range1.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range1.Tables[1].Borders.Enable = 1;
                    Paragraph paragraph3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph3.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph3.Range.Text = "\nS1:2-Currencies\n";
                    paragraph3.Range.InsertParagraphAfter();
                    paragraph3.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt).Range;
                    range3.Text = "Currency\tAbbreviation\tActive Account\tPassive Account\n";
                    OracleCommand command1 = Frm_1.dbcon.CreateCommand();
                    command1.CommandText = "select treferencecurrency.currency,treferencecurrency.abbreviation,activeaccount,passiveaccount from a4m.treferencecurrency where branch=" + Frm_1.bank_num + " and favorite=1 and activeaccount is not null and passiveaccount is not null order by 1";
                    OracleDataReader oracleDataReader1 = command1.ExecuteReader();
                    while (oracleDataReader1.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range4 = range3;
                        string str2 = range4.Text + oracleDataReader1[0].ToString() + "\t" + oracleDataReader1[1].ToString() + "\t" + oracleDataReader1[2].ToString() + "\t" + oracleDataReader1[3].ToString() + "\n";
                        range4.Text = str2;
                    }
                    oracleDataReader1.Close();
                    List_of_dictionaries.col_width = 120;
                    range3.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range3.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range3.Tables[1].Borders.Enable = 1;
                    Paragraph paragraph4 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph4.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph4.Range.Text = "\nS1:3-Account Chart\n";
                    paragraph4.Range.InsertParagraphAfter();
                    paragraph4.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range5 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt).Range;
                    range5.Text = "Ledger\tSubledger\tAccount\tName\tCurrencyno\tState\tType\n";
                    OracleCommand command2 = Frm_1.dbcon.CreateCommand();
                    command2.CommandText = "select ledger,subledger,Internal Account,replace(Name,'\n') Name,Currencyno,decode(sign,0,'Liability',1,'Active',2,'Nominal',3, 'Credit') State,decode(consolidated,1,'Consolidated',2,'Non Consolidated')Type from a4m.tplanaccount where branch =" + Frm_1.bank_num + " order by 1,2";
                    OracleDataReader oracleDataReader2 = command2.ExecuteReader();
                    while (oracleDataReader2.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range4 = range5;
                        string str2 = range4.Text + oracleDataReader2[0].ToString() + "\t" + oracleDataReader2[1].ToString() + "\t" + oracleDataReader2[2].ToString() + "\t" + oracleDataReader2[3].ToString() + "\t" + oracleDataReader2[4].ToString() + "\t" + oracleDataReader2[5].ToString() + "\t" + oracleDataReader2[6].ToString() + "\n";
                        range4.Text = str2;
                    }
                    oracleDataReader2.Close();
                    range5.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range5.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range5.Tables[1].Borders.Enable = 1;
                    Paragraph paragraph5 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph5.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph5.Range.Text = "\nS1:4-Correspondance Of TWCMS Accounting Transaction ID With Bank System ID\n";
                    paragraph5.Range.InsertParagraphAfter();
                    paragraph5.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range6 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt).Range;
                    range6.Text = "IDENT\tNAME\tORIGMSGTYPE\tMSGTYPE\tPROCCODE\tDESCRIPTION\tREVERSAL\n";
                    OracleCommand command3 = Frm_1.dbcon.CreateCommand();
                    command3.CommandText = "SELECT IDENTTWR IDENT,NAMETWR NAME,ORIGMSGTYPE,MSGTYPE,PROCCODE,DESCRIPTION,DECODE(UPPER(REVERSAL),'R','REVERSAL','N','NOT REVERSAL',REVERSAL) REVERSAL FROM A4M.TREFERENCEENTRYCTL WHERE BRANCH= " + Frm_1.bank_num + " order by 1";
                    OracleDataReader oracleDataReader3 = command3.ExecuteReader();
                    while (oracleDataReader3.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range4 = range6;
                        string str2 = range4.Text + oracleDataReader3[0].ToString() + "\t" + oracleDataReader3[1].ToString() + "\t" + oracleDataReader3[2].ToString() + "\t" + oracleDataReader3[3].ToString() + "\t" + oracleDataReader3[4].ToString() + "\t" + oracleDataReader3[5].ToString() + "\t" + oracleDataReader3[6].ToString() + "\n";
                        range4.Text = str2;
                    }
                    oracleDataReader3.Close();
                    range6.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range6.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range6.Tables[1].Borders.Enable = 1;
                    oPara1 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                }
                if (flag2)
                {
                    oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    oPara1.Range.Text = "\fSection 2 : Card Products\n";
                    oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara1.Range.Font.Bold = 0;
                    oPara1.Range.Font.Size = 9f;
                    oPara1.Format.SpaceAfter = 0.0f;
                    oPara1.Range.InsertParagraphAfter();
                    Paragraph paragraph2 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph2.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph2.Range.Text = "S2:1-(A) Main Product Characteristics\n";
                    paragraph2.Range.InsertParagraphAfter();
                    Paragraph paragraph3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph3.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range1 = paragraph3.Range;
                    range1.Text = "Product Name\tPrefix\tExpiry (Months)\tService Code\tDefault Usage Limit\tDefault Financial Profile\tState (At Issue)\tStatus (At Issue)\n";
                    List_of_dictionaries.c.CommandText = "select a4m.treferencecardproduct.NAME,prefix BIN,period Validity_Months,ServiceCODE,A4M.tlimitgrp.GRPNAME Usage_Limit,A4M.TFINPROFILE.NAME Financial_Profile,A4M.TREFERENCECARDSIGN.NAME State_upon_Creation,a4m.referencecrd_stat.getname(a4m.treferencecardproduct.Status) Status_upon_Creation from a4m.treferencecardproduct,a4m.tlimitgrp,a4m.tfinprofile,a4m.treferencecardsign where a4m.treferencecardproduct.branch=" + Frm_1.bank_num + " and a4m.tfinprofile.branch(+)=a4m.treferencecardproduct.branch and a4m.tfinprofile.profile(+)=a4m.treferencecardproduct.finprofile and a4m.treferencecardsign.branch(+)=a4m.treferencecardproduct.branch and a4m.treferencecardsign.cardsign(+)=a4m.treferencecardproduct.state and a4m.tlimitgrp.branch(+)=a4m.treferencecardproduct.branch and a4m.tlimitgrp.id(+)=a4m.treferencecardproduct.GRPLimitID order by 1";
                    OracleDataReader oracleDataReader1 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader1.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range1;
                        string str = range2.Text + oracleDataReader1[0].ToString() + "\t" + oracleDataReader1[1].ToString() + "\t" + oracleDataReader1[2].ToString() + "\t" + oracleDataReader1[3].ToString() + "\t" + oracleDataReader1[4].ToString() + "\t" + oracleDataReader1[5].ToString() + "\t" + oracleDataReader1[6].ToString() + "\t" + oracleDataReader1[7].ToString() + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader1.Close();
                    range1.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range1.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range1.Tables[1].Borders.Enable = 1;
                    Paragraph paragraph4 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph4.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph4.Range.Text = "\nS2:1-(B) Pan Generation Method\n";
                    paragraph4.Range.InsertParagraphAfter();
                    Paragraph paragraph5 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph5.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range3 = paragraph5.Range;
                    range3.Text = "Product Name\tGeneration Order\tGeneration Type\n";
                    List_of_dictionaries.c.CommandText = "select treferencecardproduct.name,code from a4m.treferencecardproduct where treferencecardproduct.branch=" + Frm_1.bank_num + " order by 1";
                    OracleDataReader oracleDataReader2 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader2.Read())
                    {
                        List_of_dictionaries.c.CommandText = "select rp.name,pr.ONUM,decode(pr.FTYPE,1,'Sequence',3,'Customer ID',2,'Card Serial No.',5,'Constant',6,'Random Number',7,'Card Sub-Branch Code (BranchPart)',8,'Customer External ID',9,'Customer Sub-Branch Code (BranchPart)',10,'Sequence With Regard To Branches',11,'Random Number With Regard To Branches',12,'Sequence (Ranges)',13,'Random Number (Ranges)',14,'PL/SQL Block',null) type,rp.code from a4m.treferencecardproduct rp,a4m.tcpcardnumber pr where rp.branch=pr.branch and rp.CODE=pr.CODE and rp.BRANCH=" + Frm_1.bank_num + " and rp.CODE=" + oracleDataReader2[1] + " order by 1";
                        OracleDataReader oracleDataReader3 = List_of_dictionaries.c.ExecuteReader();
                        string str1 = "";
                        string str2 = "";
                        while (oracleDataReader3.Read())
                        {
                            str1 = str1 + oracleDataReader3[1].ToString() + "                                      ";
                            str2 = str2 + oracleDataReader3[2].ToString() + "                                      ";
                        }
                        oracleDataReader3.Close();
                        Microsoft.Office.Interop.Word.Range range2 = range3;
                        string str3 = range2.Text + oracleDataReader2[0].ToString() + "\t" + str1 + "\t" + str2 + "\n";
                        range2.Text = str3;
                    }
                    oracleDataReader2.Close();
                    range3.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                    range3.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range3.Tables[1].Borders.Enable = 1;
                    Paragraph paragraph6 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph6.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph6.Range.Text = "\nS2:2-List of branches\n";
                    paragraph6.Range.InsertParagraphAfter();
                    Paragraph paragraph7 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph7.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range4 = paragraph7.Range;
                    range4.Text = "Branch Code\tBranch name\n";
                    List_of_dictionaries.c.CommandText = "select branchpart,name from a4m.tbranchpart where branch=" + Frm_1.bank_num + " order by 1";
                    OracleDataReader oracleDataReader4 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader4.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range4;
                        string str = range2.Text + oracleDataReader4[0].ToString() + "\t" + oracleDataReader4[1].ToString() + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader4.Close();
                    range4.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range4.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range4.Tables[1].Borders.Enable = 1;
                    Paragraph paragraph8 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph8.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph8.Range.Text = "\nS2:3-List of regions\n";
                    paragraph8.Range.InsertParagraphAfter();
                    Paragraph paragraph9 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph9.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range5 = paragraph9.Range;
                    range5.Text = "Region Code\tRegion name\n";
                    List_of_dictionaries.c.CommandText = "select code,name from a4m.treferenceregion where branch=" + Frm_1.bank_num + " order by 1";
                    OracleDataReader oracleDataReader5 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader5.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range5;
                        string str = range2.Text + oracleDataReader5[0].ToString() + "\t" + oracleDataReader5[1].ToString() + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader5.Close();
                    range5.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range5.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range5.Tables[1].Borders.Enable = 1;
                    Paragraph oPara10 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    oPara10.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    oPara10.Range.Text = "\nS2:4-List of cities\n";
                    oPara10.Range.InsertParagraphAfter();
                    Paragraph paragraph11 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph11.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range6 = paragraph11.Range;
                    range6.Text = "City Code\tCity Name\tRelated Region Code\tRelated Region Name\n";
                    List_of_dictionaries.c.CommandText = "select t1.code,t1.name,t2.code related_region_code,t2.name related_region_name from a4m.treferencecity t1,a4m.treferenceregion t2 where t1.branch=t2.branch and t1.region=t2.code and t1.branch=" + Frm_1.bank_num + " order by 1";
                    OracleDataReader oracleDataReader6 = List_of_dictionaries.c.ExecuteReader();

                    //new code
                    StringBuilder builder = new StringBuilder();
                    while (oracleDataReader6.Read())
                    {
                        builder.AppendLine(oracleDataReader6[0].ToString() + "\t" + oracleDataReader6[1].ToString() + "\t" + oracleDataReader6[2].ToString() + "\t" + oracleDataReader6[3].ToString());
                        range6.Text = builder.ToString();
                        builder.Length = 0;
                    }


                    oracleDataReader6.Close();

                    range6.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);

                    range6.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range6.Tables[1].Borders.Enable = 1;
                    oPara1 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                }
                Paragraph paragraph12;
                if (flag3)
                {
                    oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    oPara1.Range.Text = "\fSection 3 : Card Profiles";
                    oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara1.Range.Font.Bold = 0;
                    oPara1.Range.Font.Size = 9f;
                    oPara1.Format.SpaceAfter = 0.0f;
                    oPara1.Range.InsertParagraphAfter();
                    Paragraph paragraph2 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph2.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph2.Range.Text = "\nAcquirer Bins\n";
                    paragraph2.Range.InsertParagraphAfter();
                    Paragraph paragraph3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph3.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range1 = paragraph3.Range;
                    range1.Text = "Visa Auth. Bin\tVisa Acquirer Bin\tDescription\n";
                    List_of_dictionaries.c.CommandText = "select distinct Bin Auth_Bin,Bin2 Acquirer_Bin,remark Description  from a4m.texbinlist where branch=" + Frm_1.bank_num;
                    OracleDataReader oracleDataReader1 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader1.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range1;
                        string str = range2.Text + oracleDataReader1[0].ToString() + "\t" + oracleDataReader1[1].ToString() + "\t" + oracleDataReader1[2].ToString() + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader1.Close();
                    range1.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range1.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range1.Tables[1].Borders.Enable = 1;
                    List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt).Range.InsertParagraphAfter();
                    Paragraph paragraph4 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph4.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range3 = paragraph4.Range;
                    range3.Text = "Acquirer IIN\tAcquirer ID\tDescription\n";
                    List_of_dictionaries.c.CommandText = "select distinct ACQIIN ,AID,Description  from a4m.ttwiassbin where branch=" + Frm_1.bank_num;
                    OracleDataReader oracleDataReader2 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader2.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range3;
                        string str = range2.Text + oracleDataReader2[0].ToString() + "\t" + oracleDataReader2[1].ToString() + "\t" + oracleDataReader2[2].ToString() + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader2.Close();
                    range3.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range3.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range3.Tables[1].Borders.Enable = 1;
                    paragraph12 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph12 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    List_of_dictionaries.c.CommandText = "select profile,name,decode(status,1,'Active','Inactive') Status from a4m.tfinprofile where branch=" + Frm_1.bank_num + " order by 1";
                    OracleDataReader oracleDataReader3 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader3.Read())
                    {
                        Paragraph paragraph5 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        paragraph5.Range.Text = "\n" + oracleDataReader3[1].ToString() + "   [" + oracleDataReader3[2].ToString() + "]";
                        paragraph5.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                        paragraph5.Range.InsertParagraphAfter();
                        Paragraph paragraph6 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        paragraph6.Range.Text = "\nS3:1-Payments\n";
                        paragraph6.Range.InsertParagraphAfter();
                        Paragraph paragraph7 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        paragraph7.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                        Microsoft.Office.Interop.Word.Range range2 = paragraph7.Range;
                        range2.Text = "Name\tValue\tCurrency\tAccount\n";
                        //EDT-921 => EDT-849 mabouleila + samr 15/04/2015
                        //List_of_dictionaries.c.CommandText = "select p.NAME,rp.AMOUNT,rp.CURRENCY,rp.ACCOUNT,p.EXECONDITION,p.PRECONDITION,NVL(SUBSTR(SUBSTR(SOURCE,INSTR(UPPER(SOURCE),'VVALUEADD')),INSTR(SUBSTR(SOURCE,INSTR(UPPER(SOURCE),'VVALUEADD')),':=')+2,(INSTR(SUBSTR(SOURCE,INSTR(UPPER(SOURCE),'VVALUEADD')),';')-INSTR(SUBSTR(SOURCE,INSTR(UPPER(SOURCE),'VVALUEADD')),':='))-2),0.00) ANNUAL_FEE from a4m.tfinpayment p,a4m.tfinprofilepayment rp ,a4m.tsqlmember sq where p.branch=rp.branch and p.payment=rp.PAYMENT and p.branch=" + Frm_1.bank_num + " and p.BRANCH=sq.branch(+) and p.EXECONDITION=sq.CODE(+) and rp.PROFILE=" + int.Parse(oracleDataReader3[0].ToString()) + " order by 1";
                        List_of_dictionaries.c.CommandText = "select p.NAME,nvl(ppp.AMOUNT,pp.AMOUNT) AMOUNT, nvl(ppp.CURRENCY,pp.CURRENCY) CURRENCY,rp.ACCOUNT,p.EXECONDITION,p.PRECONDITION,NVL(SUBSTR(SUBSTR(SOURCE,INSTR(UPPER(SOURCE),'VVALUEADD')),INSTR(SUBSTR(SOURCE,INSTR(UPPER(SOURCE),'VVALUEADD')),':=')+2,(INSTR(SUBSTR(SOURCE,INSTR(UPPER(SOURCE),'VVALUEADD')),';')-INSTR(SUBSTR(SOURCE,INSTR(UPPER(SOURCE),'VVALUEADD')),':='))-2),0.00) ANNUAL_FEE FROM a4m.tfinpayment p, a4m.tfinprofilepayment rp , a4m.tfinplanpayment pp , a4m.tFinProfilePlanPayment ppp , a4m.tsqlmember sq where p.branch=" + Frm_1.bank_num + " and rp.PROFILE=" + int.Parse(oracleDataReader3[0].ToString()) + " AND p.branch = rp.branch AND p.payment = rp.PAYMENT AND pp.branch = p.branch AND pp.payment = p.payment AND ppp.branch(+) = rp.branch AND ppp.profile(+) = rp.profile AND ppp.payment(+) = rp.payment and p.BRANCH=sq.branch(+) and p.EXECONDITION=sq.CODE(+)  order by 1";
                        OracleDataReader oracleDataReader4 = List_of_dictionaries.c.ExecuteReader();
                        while (oracleDataReader4.Read())
                        {
                            if ((double)float.Parse(oracleDataReader4[6].ToString()) != 0.0)
                            {
                                Microsoft.Office.Interop.Word.Range range4 = range2;
                                string str = range4.Text + oracleDataReader4[0].ToString() + "\t" + oracleDataReader4[1].ToString() + " Monthly " + oracleDataReader4[6].ToString() + " Annualy\t" + oracleDataReader4[2].ToString() + "\t" + oracleDataReader4[3].ToString() + "\n";
                                range4.Text = str;
                            }
                            else
                            {
                                Microsoft.Office.Interop.Word.Range range4 = range2;
                                string str = range4.Text + oracleDataReader4[0].ToString() + "\t" + oracleDataReader4[1].ToString() + "\t" + oracleDataReader4[2].ToString() + "\t" + oracleDataReader4[3].ToString() + "\n";
                                range4.Text = str;
                            }
                        }
                        oracleDataReader4.Close();
                        range2.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                        range2.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                        range2.Tables[1].Borders.Enable = 1;
                        //Islam Atta

                        Paragraph paragraph10 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        paragraph10.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                        paragraph10.Range.Text = "\nS3:3-Additional Payment\n ";
                        paragraph10.Range.InsertParagraphAfter();
                        Paragraph paragraph010 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        paragraph010.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                        Microsoft.Office.Interop.Word.Range range9 = paragraph010.Range;
                        range9.Text = "Name\tPercentage\tAmount\tMinAmount\tMaxAmount\tCurrency\tAccount\n";
                        //                        List_of_dictionaries.c.CommandText = @"
                        //  SELECT tfc.name,
                        //         tfpc2.PRC,
                        //         tfpc2.AMOUNT,
                        //         TFP.MINAMOUNT,
                        //         TFP.MAXAMOUNT,
                        //         TFP.CURRENCY,
                        //         TFC.ACCOUNT
                        //    FROM  a4m.tfinprofileplancommission tfpc1,
                        //         a4m.tfinprofileplancommission tfpc2,
                        //         a4m.tfinplancommission tfp,
                        //         a4m.tfincommission tfc

                        //   WHERE     tfpc1.branch = " + Frm_1.bank_num + @"
                        //         AND tfpc1.BRANCH = tfpc2.BRANCH
                        //         AND tfpc1.COMMISSION = tfpc2.PARENTPAYMENT
                        //         AND tfpc1.profile =" + int.Parse(oracleDataReader3[0].ToString()) + @"
                        //         AND tfpc1.profile = tfpc2.PROFILE
                        //         AND tfp.branch = tfpc2.branch
                        //         AND tfp.COMMISSION = tfpc2.COMMISSION
                        //         AND tfc.branch = tfpc2.branch
                        //         AND TFC.COMMISSION = tfpc2.COMMISSION
                        //ORDER BY 1";

                        //msattar SBP-9547 -> 14-07-2020
                        List_of_dictionaries.c.CommandText = @"
 --additional payments -> in cms its under payments under addtional fees undeer tariffs
--and we note that the payment tab related to two tables tfinprofilepaymnet and tfinprofileplanpaymnet
  SELECT tfc.name,
         tfpc1.PRC,
         tfpc1.AMOUNT,
         TFP.MINAMOUNT,
         TFP.MAXAMOUNT,
         TFP.CURRENCY,
         TFC.ACCOUNT
    FROM  a4m.tfinprofileplancommission tfpc1,
    a4m.tfinprofileplanpayment tfpp2,
         --a4m.tfinprofileplancommission tfpc2,
         a4m.tfinplancommission tfp,
         a4m.tfincommission tfc
        
   WHERE     tfpc1.branch = " + Frm_1.bank_num + @"
         --AND tfpc1.BRANCH = tfpc2.BRANCH
         AND tfpc1.BRANCH = tfpp2.BRANCH
         --AND tfpc1.COMMISSION <> tfpc2.PARENTPAYMENT
         AND tfpc1.PARENTPAYMENT = tfpp2.payment
         AND tfpc1.profile =" + int.Parse(oracleDataReader3[0].ToString()) + @"
         --AND tfpc1.profile = tfpc2.PROFILE
         AND tfpc1.profile = tfpp2.PROFILE
         AND tfp.branch = tfpc1.branch
         AND tfp.COMMISSION = tfpc1.COMMISSION
         AND tfc.branch = tfpc1.branch
         AND TFC.COMMISSION = tfpc1.COMMISSION
         
         union 
         --additional payments -> in cms its under payments under addtional fees undeer tariffs
--and we note that the payment tab related to two tables tfinprofilepaymnet and tfinprofileplanpaymnet
  SELECT tfc.name,
         tfp.PRC,
         tfp.AMOUNT,
         TFP.MINAMOUNT,
         TFP.MAXAMOUNT,
         TFP.CURRENCY,
         TFC.ACCOUNT
    FROM  a4m.tfinprofilecommission tfpc1,
    a4m.tfinprofilepayment tfpp2,
         --a4m.tfinprofileplancommission tfpc2,
         a4m.tfinplancommission tfp,
         a4m.tfincommission tfc
        
   WHERE     tfpc1.branch = " + Frm_1.bank_num + @" --58
         --AND tfpc1.BRANCH = tfpc2.BRANCH
         AND tfpc1.BRANCH = tfpp2.BRANCH
         --AND tfpc1.COMMISSION <> tfpc2.PARENTPAYMENT
         AND tfpc1.PARENTPAYMENT = tfpp2.payment
         AND tfpc1.profile = " + int.Parse(oracleDataReader3[0].ToString()) + @" --11
         --AND tfpc1.profile = tfpc2.PROFILE
         AND tfpc1.profile = tfpp2.PROFILE
         AND tfp.branch = tfpc1.branch
         AND tfp.COMMISSION = tfpc1.COMMISSION
         AND tfc.branch = tfpc1.branch
         AND TFC.COMMISSION = tfpc1.COMMISSION
         and not exists (
         select * from a4m.tfinprofileplancommission
         where branch = tfpc1.branch
         and profile = tfpc1.profile
         and commission = tfpc1.commission
         and parentpayment = tfpc1.parentpayment)
         
ORDER BY 1";
                        OracleDataReader oracleDataReader7 = List_of_dictionaries.c.ExecuteReader();
                        while (oracleDataReader7.Read())
                        {
                            Microsoft.Office.Interop.Word.Range range4 = range9;
                            string str = range4.Text + oracleDataReader7[0].ToString() + "\t" + oracleDataReader7[1].ToString() + "\t" + oracleDataReader7[2].ToString() + "\t" + oracleDataReader7[3].ToString() + "\t" + oracleDataReader7[4].ToString() + "\t" + oracleDataReader7[5].ToString() + "\t" + oracleDataReader7[6].ToString() + "\n";
                            range4.Text = str;
                        }
                        oracleDataReader7.Close();
                        range9.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                        range9.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                        range9.Tables[1].Borders.Enable = 1;

                        Paragraph paragraph8 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        paragraph8.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                        paragraph8.Range.Text = "\nS3:2-Fees\n ";
                        paragraph8.Range.InsertParagraphAfter();
                        Paragraph paragraph9 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        paragraph9.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                        Microsoft.Office.Interop.Word.Range range5 = paragraph9.Range;
                        range5.Text = "Name\tPercentage\tAmount\tMinAmount\tMaxAmount\tCurrency\tAccount\n";
                        //   List_of_dictionaries.c.CommandText = "select p.NAME,nvl(rp.prc,0)prc,nvl(rp.AMOUNT,0) AMOUNT,rp.Minamount,rp.maxamount,nvl(rp.currency,rp2.CURRENCY) currency,rp2.account from a4m.tfincommission p,a4m.tfinprofileplancommission rp ,a4m.tfinprofilecommission rp2 where p.branch=rp2.branch and p.commission=rp2.commission and rp.branch(+) =rp2.branch and rp.profile(+)=rp2.profile and rp.commission(+) =rp2.commission and p.branch= " + Frm_1.bank_num + "  and rp2.PROFILE=" + int.Parse(oracleDataReader3[0].ToString()) + " order by 1";
                        //Islam Atta 
                        //old
                        /*List_of_dictionaries.c.CommandText = @"SELECT p.NAME,NVL (rp.prc, 0) prc,NVL (rp.AMOUNT, 0) AMOUNT,rp.Minamount,rp.maxamount,NVL (rp.currency, rp2.CURRENCY) currency,rp2.account FROM a4m.tfincommission p,a4m.tfinprofileplancommission rp,a4m.tfinprofilecommission rp2 WHERE     p.branch = rp2.branch AND p.commission = rp2.commission AND rp.branch(+) = rp2.branch AND rp.profile(+) = rp2.profile AND rp.commission(+) = rp2.commission AND p.branch =" + Frm_1.bank_num + "AND rp2.PROFILE =" + int.Parse(oracleDataReader3[0].ToString()) + @" 
                        AND p.commission Not In  (SELECT TF2.commission FROM a4m.tfinprofilecommission tf1,a4m.tfinprofilecommission tf2,a4m.tfinplancommission tfp,a4m.tfincommission tfc WHERE     tf1.branch =" + Frm_1.bank_num + " AND TF1.BRANCH = TF2.BRANCH AND TF1.COMMISSION = TF2.PARENTCOMMISSION AND tf1.profile =" + int.Parse(oracleDataReader3[0].ToString()) + @" AND tf1.profile = TF2.PROFILE AND tfp.branch = tf2.branch AND tfp.COMMISSION = tf2.COMMISSION AND tfc.branch = tf2.branch AND TFC.COMMISSION = TF2.COMMISSION
                        union
                        SELECT tfpc2.COMMISSION
FROM  a4m.tfinprofileplancommission tfpc1,a4m.tfinprofileplancommission tfpc2,a4m.tfinplancommission tfp,a4m.tfincommission tfc 
WHERE     tfpc1.branch = " + Frm_1.bank_num + @" AND tfpc1.BRANCH = tfpc2.BRANCH
AND tfpc1.COMMISSION = tfpc2.PARENTPAYMENT
AND tfpc1.profile =" + int.Parse(oracleDataReader3[0].ToString()) + @"
AND tfpc1.profile = tfpc2.PROFILE
AND tfp.branch = tfpc2.branch
AND tfp.COMMISSION = tfpc2.COMMISSION
AND tfc.branch = tfpc2.branch
AND TFC.COMMISSION = tfpc2.COMMISSION) order by 1";*/

                        // + Frm_1.bank_num + @" --58
                        //" + int.Parse(oracleDataReader3[0].ToString()) + @" --11
                        //new msattar SBP-9547
                        List_of_dictionaries.c.CommandText = @"   --THE SQL SAY GET VALUES FROM REDIFIND UNION FROM DICTIONRAY BUT WE ADD CONDTION ON PART OF FROM DICTIONARY THAT SAY IF COUNT == 0 THEN THERE IS NO REDIFINED
 --VALUES SO DICTIONARY OF THE SECOND PART OF UNION WILL GET THE VALUES, IF COUNTS>0 THE THE FIRST PART OF UNION WILL GET VALUES MEANS THE REDIFIND ONE
  SELECT  p.NAME namee,
          CASE  RP2.INHERITFLAG WHEN '1111111111' THEN plancom.prc ELSE rp.prc END  prc, --msattar add to nvl plancom.prc
          CASE  RP2.INHERITFLAG WHEN '1111111111' THEN plancom.AMOUNT ELSE rp.AMOUNT END  AMOUNT, --msattar add to nvl plancom.amount
          CASE  RP2.INHERITFLAG WHEN '1111111111' THEN plancom.Minamount ELSE rp.Minamount END  Minamount, --msattar add nvl
          CASE  RP2.INHERITFLAG WHEN '1111111111' THEN plancom.MAXAMOUNT ELSE rp.MAXAMOUNT END  maxamount, --msattar add nvl
          CASE  RP2.INHERITFLAG WHEN '1111111111' THEN plancom.currency ELSE rp.currency END  currency,
          P.account,
          RP2.INHERITFLAG
    FROM a4m.tfincommission p, --TO GET THE NAME
         a4m.tfinprofileplancommission rp, --THIS HAVE THE REDIFIND VALUES.AND IF U PICK FROM DICTIONARY THE EQUEVELNT COLOMN
         -- HERE WILL BE DELETED FROM REDIFIND TAP IN CMS AND HERE FROM TFINPROFILEPALNCOMMSISSION
         a4m.tfinprofilecommission rp2,
         a4m.tFinPlanCommission plancom --msattar THISHAVE THE FROM DICTIONARY VALUES.
   WHERE   
             
          p.branch = rp2.branch
         AND p.commission = rp2.commission
         AND rp.branch(+) = rp2.branch
         AND rp.profile(+) = rp2.profile
         AND rp.commission(+) = rp2.commission
         and rp2.COMMISSION = plancom.COMMISSION(+)  --sattar
         AND rp2.BRANCH = plancom.BRANCH(+)  --msattar
         AND p.branch = " + Frm_1.bank_num + @"--3
         AND rp2.PROFILE = " + int.Parse(oracleDataReader3[0].ToString()) + @" --11
         and rp2.PARENTCOMMISSION =0  --msattar
         and rp2.PARENTPAYMENT =0     --msattar
         AND p.commission NOT IN
                (SELECT TF2.commission
                   FROM a4m.tfinprofilecommission tf1,
                        a4m.tfinprofilecommission tf2,
                        a4m.tfinplancommission tfp,
                        a4m.tfincommission tfc
                  WHERE     tf1.branch = " + Frm_1.bank_num + @"--3
                        AND TF1.BRANCH = TF2.BRANCH
                        AND TF1.COMMISSION = TF2.PARENTCOMMISSION
                        AND tf1.profile =   " + int.Parse(oracleDataReader3[0].ToString()) + @" --11
                        AND tf1.profile = TF2.PROFILE
                        AND tfp.branch = tf2.branch
                        AND tfp.COMMISSION = tf2.COMMISSION
                        AND tfc.branch = tf2.branch
                        AND TFC.COMMISSION = TF2.COMMISSION
                 UNION
                 SELECT tfpc2.COMMISSION
                   FROM a4m.tfinprofileplancommission tfpc1,
                        a4m.tfinprofileplancommission tfpc2,
                        a4m.tfinplancommission tfp,
                        a4m.tfincommission tfc
                  WHERE     tfpc1.branch = " + Frm_1.bank_num + @"--3
                        AND tfpc1.BRANCH = tfpc2.BRANCH
                        AND tfpc1.COMMISSION = tfpc2.PARENTPAYMENT
                        AND tfpc1.profile =  " + int.Parse(oracleDataReader3[0].ToString()) + @" --11
                        AND tfpc1.profile = tfpc2.PROFILE
                        AND tfp.branch = tfpc2.branch
                        AND tfp.COMMISSION = tfpc2.COMMISSION
                        AND tfc.branch = tfpc2.branch
                        AND TFC.COMMISSION = tfpc2.COMMISSION)   
 
ORDER BY 1";

                        OracleDataReader oracleDataReader5 = List_of_dictionaries.c.ExecuteReader();
                        while (oracleDataReader5.Read())
                        {
                            Microsoft.Office.Interop.Word.Range range4 = range5;
                            string str = range4.Text + oracleDataReader5[0].ToString() + "\t" + oracleDataReader5[1].ToString() + "\t" + oracleDataReader5[2].ToString() + "\t" + oracleDataReader5[3].ToString() + "\t" + oracleDataReader5[4].ToString() + "\t" + oracleDataReader5[5].ToString() + "\t" + oracleDataReader5[6].ToString() + "\n";
                            range4.Text = str;
                        }
                        oracleDataReader5.Close();
                        //oracleDataReader4.Close();
                        range5.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                        range5.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                        range5.Tables[1].Borders.Enable = 1;

                        //islam atta

                        Paragraph paragraph09 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        paragraph09.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                        paragraph09.Range.Text = "\nS3:3-Additional Fees\n ";
                        paragraph09.Range.InsertParagraphAfter();
                        Paragraph paragraph009 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        paragraph009.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                        Microsoft.Office.Interop.Word.Range range6 = paragraph009.Range;
                        range6.Text = "Name\tPercentage\tAmount\tMinAmount\tMaxAmount\tCurrency\tAccount\n";
                        //old
                        //List_of_dictionaries.c.CommandText = "SELECT tfc.name,TFC.PRC,tfp.AMOUNT,TFP.MINAMOUNT,TFP.MAXAMOUNT,TFP.CURRENCY,TFC.ACCOUNT FROM a4m.tfinprofilecommission tf1,a4m.tfinprofilecommission tf2,a4m.tfinplancommission tfp,a4m.tfincommission tfc WHERE tf1.branch =" + Frm_1.bank_num + "AND TF1.BRANCH = TF2.BRANCH AND TF1.COMMISSION = TF2.PARENTCOMMISSION AND tf1.profile =" + int.Parse(oracleDataReader3[0].ToString()) + " AND tf1.profile = TF2.PROFILE AND tfp.branch = tf2.branch AND tfp.COMMISSION = tf2.COMMISSION AND tfc.branch = tf2.branch AND TFC.COMMISSION = TF2.COMMISSION  order by 1";

                        //new msattar SBP-9547
                        List_of_dictionaries.c.CommandText = @"/* Formatted on 6/22/2020 2:33:13 PM (QP5 v5.227.12220.39754) */
SELECT tfc.name,
         tf2.PRC,--tf
         tfp.AMOUNT,
         TFP.MINAMOUNT,
         TFP.MAXAMOUNT,
         TFP.CURRENCY,
         TFC.ACCOUNT, TF2.PARENTCOMMISSION
    FROM a4m.tfinprofilecommission tf1,
         a4m.tfinprofileplancommission tf2,
         a4m.tfinplancommission tfp,
         a4m.tfincommission tfc
         --a4m.tfinprofileplancommission
   WHERE     tf1.branch = " + Frm_1.bank_num + @" --7
         AND TF1.BRANCH = TF2.BRANCH
          and TF1.COMMISSION = TF2.PARENTCOMMISSION
         AND tf1.profile = " + int.Parse(oracleDataReader3[0].ToString()) + @" --20
         AND tf1.profile = TF2.PROFILE
         AND tfp.branch = tf2.branch
         AND tfp.COMMISSION = tf2.COMMISSION
         AND tfc.branch = tf2.branch
         AND TFC.COMMISSION = TF2.COMMISSION
         
         union
         
  SELECT tfc.name,
         tfp.PRC,--tf
         tfp.AMOUNT,
         TFP.MINAMOUNT,
         TFP.MAXAMOUNT,
         TFP.CURRENCY,
         TFC.ACCOUNT, TF2.PARENTCOMMISSION
    FROM a4m.tfinprofilecommission tf1,
         a4m.tfinprofilecommission tf2,
         a4m.tfinplancommission tfp,
         a4m.tfincommission tfc
         --a4m.tfinprofileplancommission
   WHERE     tf1.branch = " + Frm_1.bank_num + @"
         AND TF1.BRANCH = TF2.BRANCH
          and TF1.COMMISSION = TF2.PARENTCOMMISSION
         AND tf1.profile = " + int.Parse(oracleDataReader3[0].ToString()) + @"
         AND tf1.profile = TF2.PROFILE
         AND tfp.branch = tf2.branch
         AND tfp.COMMISSION = tf2.COMMISSION
         AND tfc.branch = tf2.branch
         AND TFC.COMMISSION = TF2.COMMISSION
         and not exists (
         select * from a4m.tfinprofileplancommission
         where branch = tf2.branch
         and profile = tf2.profile
         and commission = tf2.commission
         and PARENTCOMMISSION = tf2.PARENTCOMMISSION)
ORDER BY 1";

                        //List_of_dictionaries.c.CommandText = "SELECT TF1.PROFILE,tf2.COMMISSION,tfc.name,tfp.AMOUNT FROM a4m.tfinprofilecommission tf1,a4m.tfinprofilecommission tf2,a4m.tfinplancommission tfp,a4m.tfincommission tfc WHERE tf1.branch = " + Frm_1.bank_num + " AND TF1.BRANCH = TF2.BRANCH AND TF1.COMMISSION = TF2.PARENTCOMMISSION AND tf1.profile =" + int.Parse(oracleDataReader3[0].ToString()) + "AND tf1.profile = TF2.PROFILE AND tfp.branch = tf2.branch AND tfp.COMMISSION = tf2.COMMISSION AND tfc.branch = tf2.branch AND TFC.COMMISSION = TF2.COMMISSION order by 1";
                        OracleDataReader oracleDataReader6 = List_of_dictionaries.c.ExecuteReader();
                        while (oracleDataReader6.Read())
                        {
                            Microsoft.Office.Interop.Word.Range range4 = range6;
                            string str = range4.Text + oracleDataReader6[0].ToString() + "\t" + oracleDataReader6[1].ToString() + "\t" + oracleDataReader6[2].ToString() + "\t" + oracleDataReader6[3].ToString() + "\t" + oracleDataReader6[4].ToString() + "\t" + oracleDataReader6[5].ToString() + "\t" + oracleDataReader6[6].ToString() + "\n";
                            range4.Text = str;
                        }
                        oracleDataReader6.Close();
                        range6.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                        range6.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                        range6.Tables[1].Borders.Enable = 1;
                    }
                    oracleDataReader3.Close();
                    oPara1 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                }
                if (flag4)
                {
                    oPara1.Range.Text = "\fSection 4 : Periodic Usage Limit Note: 999 & 999999 means no limitation";
                    oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara1.Range.Font.Bold = 0;
                    oPara1.Range.Font.Size = 9f;
                    oPara1.Format.SpaceAfter = 0.0f;
                    oPara1.Range.InsertParagraphAfter();
                    List_of_dictionaries.c.CommandText = "select grpname,id from a4m.tlimitgrp where branch=" + Frm_1.bank_num + " order by 1";
                    OracleDataReader oracleDataReader1 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader1.Read())
                    {
                        Paragraph paragraph2 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        paragraph2.Range.Text = "\n" + oracleDataReader1[0].ToString() + "\n";
                        paragraph2.Range.InsertParagraphAfter();
                        Paragraph paragraph3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        paragraph3.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                        Microsoft.Office.Interop.Word.Range range1 = paragraph3.Range;
                        range1.Text = "Limit Name\tLimit ID\tLimit Value\tReset Period\tReset On\n";
                        if (Frm_1.bank_num == "69")
                        {
                            List_of_dictionaries.c.CommandText = "select treferencecardlimit.name,tlimititem.LIMITID,tlimititem.MAXVALUE,DECODE (tlimititem.PERIODTYPE, 0, 'Day', 1, 'One Week', 2, 'Month', 3, 'Quarter', 4, 'Year', 5, 'Calendar date', 6, 'Infinite', 7, 'Sliding', 8, 'Single', 9, 'Reset at refresh', 12, 'Daily', 'Not Defined'),((tlimititem.period/60)/24) from a4m.tlimititem,a4m.treferencecardlimit where tlimititem.branch=treferencecardlimit.branch and tlimititem.LIMITID=treferencecardlimit.LIMITID and tlimititem.branch=" + Frm_1.bank_num + " and tlimititem.LIMITGRPID=" + int.Parse(oracleDataReader1[1].ToString()) + " order by 1";
                        }
                        else if (Frm_1.bank_num == "73")
                        {
                            List_of_dictionaries.c.CommandText = "select treferencecardlimit.name,tlimititem.LIMITID,tlimititem.MAXVALUE,DECODE (tlimititem.PERIODTYPE, 0, 'Day', 1, 'One Week', 2, 'Month', 3, 'Quarter', 4, 'Year', 5, 'Calendar date', 6, 'Infinite', 7, 'Sliding', 8, 'Single', 9, 'Reset at refresh', 12, 'Daily', 'Not Defined'),((tlimititem.period/60)/24) from a4m.tlimititem,a4m.treferencecardlimit where tlimititem.branch=treferencecardlimit.branch and tlimititem.LIMITID=treferencecardlimit.LIMITID and tlimititem.branch=" + Frm_1.bank_num + " and tlimititem.LIMITGRPID=" + int.Parse(oracleDataReader1[1].ToString()) + " order by 1";
                        }
                        else
                        {
                            List_of_dictionaries.c.CommandText = "select treferencecardlimit.name,tlimititem.LIMITID,tlimititem.MAXVALUE,decode(tlimititem.PERIODTYPE,0,'Daily',1,'Weekly',2,'Monthly',3,'Quarterly',7,'Yearly',4,'Infinite',5,'Single',6,'Reset at refresh',null,' ','Not Defined'),((tlimititem.period/60)/24) from a4m.tlimititem,a4m.treferencecardlimit where tlimititem.branch=treferencecardlimit.branch and tlimititem.LIMITID=treferencecardlimit.LIMITID and tlimititem.branch=" + Frm_1.bank_num + " and tlimititem.LIMITGRPID=" + int.Parse(oracleDataReader1[1].ToString()) + " order by 1";
                        }

                        OracleDataReader oracleDataReader2 = List_of_dictionaries.c.ExecuteReader();
                        while (oracleDataReader2.Read())
                        {
                            Microsoft.Office.Interop.Word.Range range2 = range1;
                            string str = range2.Text + oracleDataReader2[0].ToString() + "\t" + oracleDataReader2[1].ToString() + "\t" + oracleDataReader2[2].ToString() + "\t" + oracleDataReader2[3].ToString() + "\t" + oracleDataReader2[4].ToString() + "\n";
                            range2.Text = str;
                        }
                        oracleDataReader2.Close();
                        range1.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                        range1.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                        range1.Tables[1].Borders.Enable = 1;
                    }
                    oracleDataReader1.Close();
                    oPara1 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                }
                if (flag5)
                {
                    oPara1.Range.Text = "\fSection 5 : Delinquency Settings - Credit Cards only";
                    oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara1.Range.Font.Bold = 0;
                    oPara1.Range.Font.Size = 9f;
                    oPara1.Format.SpaceAfter = 0.0f;
                    oPara1.Range.InsertParagraphAfter();
                    List_of_dictionaries.c.CommandText = "select type Code,Name  from a4m.tcontracttype where branch=" + Frm_1.bank_num + " and status='1' order by 1";
                    OracleDataReader oracleDataReader1 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader1.Read())
                    {
                        oPara1.Range.Text = "\nContract Type: " + oracleDataReader1[1] + "\n";
                        oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                        oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        oPara1.Range.Font.Bold = 0;
                        oPara1.Range.Font.Size = 9f;
                        oPara1.Format.SpaceAfter = 0.0f;
                        oPara1.Range.InsertParagraphAfter();
                        oPara1 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        oPara1.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                        Microsoft.Office.Interop.Word.Range range1 = oPara1.Range;
                        range1.Text = "ID - Method \tDescription\tOD Days\tOL %\tEffect to Card Status\tEffect to Blocking Card Reissue\tEffect to Account Status\tEffect to Blocking Contract\tGenerate STMT\tUse Allowed OL\tContract Stick State\tCharge Interest\tCharge OD Fee\tCharge OL Fee\tCharge Service Fee\tStick Card Status\tCharge Credit Shield\n";
                        //List_of_dictionaries.c.CommandText = "select t2.STATECODE,t2.STATENAME,to_char(t1.period) overduedays,to_char(t1.Overlimit) Overlimit ,t4.NAME,decode(t2.REISSUEBAN,0,'No',1,'Yes'),decode(t2.Accblock,1,'Open',2,'Credit Only',3,'Primary only',5,'View Only',9,'Closed'),decode(t2.CONTRACTBLOCK,0,'No',1,'Yes',' '),decode(t2.STATEMENTGEN,0,'No',1,'Yes',' '),decode(t2.USEALLOWEDOL,0,'No',1,'Yes','No'),decode(t2.STICKSTATE,0,'Do not Stick',1,'Stick',2,'Stick Below',3,'Stick Above',' '),decode(t2.CHARGEINT,0,'Do not Charge',1,'Charge',2,'Suspend',3,'Accumulate',' '),decode(t2.CHARGEOVDFEE,0,'No',1,'Yes',' '),decode(t2.CHARGEOVLFEE,0,'No',1,'Yes',' '),decode(t2.SERVICEFEE,2,'Do not Charge',1,'Charge',0,'Suspend',' ') Charge_Fee,DECODE (t2.STICKCARDSTATUS,  0, 'No',  1, 'Yes',  'No') STICK_CARD_STATUS ,DECODE (t2.CHARGECRDSHIELD,  0, 'No',  1, 'Yes',  'No') CHARGE_CREDIT_SHIELD,'AUTO' method,t2.sortorder from a4m.tcontractdelinqsetup t1,a4m.tcontractstatereference t2,a4m.tcontracttype t3 ,a4m.treferencecrd_stat t4 where t2.branch=" + Frm_1.bank_num + "and t1.branch=t2.branch and t1.stateid=t2.stateid and t1.branch=t3.branch and T1.CONTRACTTYPE=T3.TYPE and period >=0  and overlimit >=0 and  t2.CARDBLOCK=t4.CRD_STAT and contracttype=" + oracleDataReader1[0] + " union all select t2.STATECODE,t2.STATENAME,'Manual' overduedays,'Manual' Overlimit ,t4.NAME,decode(t2.REISSUEBAN,0,'No',1,'Yes'),decode(t2.Accblock,1,'Open',2,'Credit Only',3,'Primary only',5,'View Only',9,'Closed'),decode(t2.CONTRACTBLOCK,0,'No',1,'Yes',' '),decode(t2.STATEMENTGEN,0,'No',1,'Yes',' '),decode(t2.USEALLOWEDOL,0,'No',1,'Yes','No'),decode(t2.STICKSTATE,0,'Do not Stick',1,'Stick',2,'Stick Below',3,'Stick Above',' '),decode(t2.CHARGEINT,0,'Do not Charge',1,'Charge',2,'Suspend',3,'Accumulate',' '),decode(t2.CHARGEOVDFEE,0,'No',1,'Yes',' '),decode(t2.CHARGEOVLFEE,0,'No',1,'Yes',' '),decode(t2.SERVICEFEE,2,'Do not Charge',1,'Charge',0,'Suspend',' ') Charge_Fee,DECODE (t2.STICKCARDSTATUS,  0, 'No',  1, 'Yes',  'No') STICK_CARD_STATUS ,DECODE (t2.CHARGECRDSHIELD,  0, 'No',  1, 'Yes',  'No') CHARGE_CREDIT_SHIELD,'Manual' method,t2.sortorder from a4m.tcontractstatereference t2 ,a4m.treferencecrd_stat t4 where t2.branch=" + Frm_1.bank_num + " and  t2.CARDBLOCK=t4.CRD_STAT and nvl((select 1 from a4m.tcontractdelinqsetup where branch=" + Frm_1.bank_num + " and stateid =t2.stateid and rownum=1 ),0)<> 1  and upper(statecode) <> 'AUTO' order by sortorder";

                        //SBN-8914  9-6-2020 msattar: change condition of overlimit to be overlimit >= -0.75 and overlimit <> 0 and thats to make sure all contracttypyes
                        //that have 'NO' in cms in deliqnuncy state in limit come with me in the query results.  
                        List_of_dictionaries.c.CommandText = "select t2.STATECODE,t2.STATENAME,to_char(t1.period) overduedays,to_char(t1.Overlimit) Overlimit ,t4.NAME,decode(t2.REISSUEBAN,0,'No',1,'Yes'),decode(t2.Accblock,1,'Open',2,'Credit Only',3,'Primary only',5,'View Only',9,'Closed'),decode(t2.CONTRACTBLOCK,0,'No',1,'Yes',' '),decode(t2.STATEMENTGEN,0,'No',1,'Yes',' '),decode(t2.USEALLOWEDOL,0,'No',1,'Yes','No'),decode(t2.STICKSTATE,0,'Do not Stick',1,'Stick',2,'Stick Below',3,'Stick Above',' '),decode(t2.CHARGEINT,0,'Do not Charge',1,'Charge',2,'Suspend',3,'Accumulate',' '),decode(t2.CHARGEOVDFEE,0,'No',1,'Yes',' '),decode(t2.CHARGEOVLFEE,0,'No',1,'Yes',' '),decode(t2.SERVICEFEE,2,'Do not Charge',1,'Charge',0,'Suspend',' ') Charge_Fee,DECODE (t2.STICKCARDSTATUS,  0, 'No',  1, 'Yes',  'No') STICK_CARD_STATUS ,DECODE (t2.CHARGECRDSHIELD,  0, 'No',  1, 'Yes',  'No') CHARGE_CREDIT_SHIELD,'AUTO' method,t2.sortorder from a4m.tcontractdelinqsetup t1,a4m.tcontractstatereference t2,a4m.tcontracttype t3 ,a4m.treferencecrd_stat t4 where t2.branch=" + Frm_1.bank_num + "and t1.branch=t2.branch and t1.stateid=t2.stateid and t1.branch=t3.branch and T1.CONTRACTTYPE=T3.TYPE and period >=0  and overlimit >= -0.75 and overlimit <> 0 and  t2.CARDBLOCK=t4.CRD_STAT and contracttype=" + oracleDataReader1[0] + " union all select t2.STATECODE,t2.STATENAME,'Manual' overduedays,'Manual' Overlimit ,t4.NAME,decode(t2.REISSUEBAN,0,'No',1,'Yes'),decode(t2.Accblock,1,'Open',2,'Credit Only',3,'Primary only',5,'View Only',9,'Closed'),decode(t2.CONTRACTBLOCK,0,'No',1,'Yes',' '),decode(t2.STATEMENTGEN,0,'No',1,'Yes',' '),decode(t2.USEALLOWEDOL,0,'No',1,'Yes','No'),decode(t2.STICKSTATE,0,'Do not Stick',1,'Stick',2,'Stick Below',3,'Stick Above',' '),decode(t2.CHARGEINT,0,'Do not Charge',1,'Charge',2,'Suspend',3,'Accumulate',' '),decode(t2.CHARGEOVDFEE,0,'No',1,'Yes',' '),decode(t2.CHARGEOVLFEE,0,'No',1,'Yes',' '),decode(t2.SERVICEFEE,2,'Do not Charge',1,'Charge',0,'Suspend',' ') Charge_Fee,DECODE (t2.STICKCARDSTATUS,  0, 'No',  1, 'Yes',  'No') STICK_CARD_STATUS ,DECODE (t2.CHARGECRDSHIELD,  0, 'No',  1, 'Yes',  'No') CHARGE_CREDIT_SHIELD,'Manual' method,t2.sortorder from a4m.tcontractstatereference t2 ,a4m.treferencecrd_stat t4 where t2.branch=" + Frm_1.bank_num + " and  t2.CARDBLOCK=t4.CRD_STAT and nvl((select 1 from a4m.tcontractdelinqsetup where branch=" + Frm_1.bank_num + " and stateid =t2.stateid and rownum=1 ),0)<> 1  and upper(statecode) <> 'AUTO' order by sortorder";

                        OracleDataReader oracleDataReader2 = List_of_dictionaries.c.ExecuteReader();
                        while (oracleDataReader2.Read())
                        {
                            Microsoft.Office.Interop.Word.Range range2 = range1;
                            string str = range2.Text + oracleDataReader2[0] + " - " + oracleDataReader2[17] + "\t" + oracleDataReader2[1] + "\t" + oracleDataReader2[2] + "\t" + oracleDataReader2[3] + "\t" + oracleDataReader2[4] + "\t" + oracleDataReader2[5] + "\t" + oracleDataReader2[6] + "\t" + oracleDataReader2[7] + "\t" + oracleDataReader2[8] + "\t" + oracleDataReader2[9] + "\t" + oracleDataReader2[10] + "\t" + oracleDataReader2[11] + "\t" + oracleDataReader2[12] + "\t" + oracleDataReader2[13] + "\t" + oracleDataReader2[14] + "\t" + oracleDataReader2[15] + "\t" + oracleDataReader2[16] + "\n";
                            range2.Text = str;
                        }
                        oracleDataReader2.Close();
                        range1.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                        range1.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                        range1.Tables[1].Borders.Enable = 1;
                    }
                    oracleDataReader1.Close();
                }
                if (flag6)
                {
                    oPara1.Range.Text = "\fSection 6 : Working Calender and Billing Cycle Calender\n";
                    oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara1.Range.Font.Bold = 0;
                    oPara1.Range.Font.Size = 9f;
                    oPara1.Format.SpaceAfter = 0.0f;
                    oPara1.Range.InsertParagraphAfter();
                    Paragraph paragraph2 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph2.Range.Text = "S6:1-Calender\n";
                    paragraph2.Range.InsertParagraphAfter();
                    Paragraph paragraph3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph3.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range1 = paragraph3.Range;
                    range1.Text = "Calender ID\tCalender Name\tCalender Days Sample\n";
                    List_of_dictionaries.c.CommandText = "select ID,Name from a4m.treferencecalendar where branch=" + Frm_1.bank_num + " order by 1";
                    OracleDataReader oracleDataReader1 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader1.Read())
                    {
                        OracleCommand command = Frm_1.dbcon.CreateCommand();
                        command.CommandText = "select distinct c.name,to_char(d.DAYS,'DAY'),count (to_char(d.DAYS,'DAY')) from a4m.treferenceholidays d,a4m.treferencecalendar c where d.branch=" + Frm_1.bank_num + " and d.BRANCH=c.BRANCH and d.ID_CLND=c.ID and c.ID=" + oracleDataReader1[0].ToString() + " group by c.name,to_char(d.DAYS,'DAY') having count (to_char(d.DAYS,'DAY')) > 20";
                        OracleDataReader oracleDataReader2 = command.ExecuteReader();
                        string str1 = "";
                        while (oracleDataReader2.Read())
                            str1 = str1 + oracleDataReader2[1].ToString() + " ";
                        oracleDataReader2.Close();
                        Microsoft.Office.Interop.Word.Range range2 = range1;
                        string str2 = range2.Text + oracleDataReader1[0] + "\t" + oracleDataReader1[1] + "\t" + str1 + "\n";
                        range2.Text = str2;
                    }
                    oracleDataReader1.Close();
                    range1.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range1.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range1.Tables[1].Borders.Enable = 1;
                    Paragraph paragraph4 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph4.Range.Text = "\nS6:2-Billing Cycle - Credit Cards Only \n";
                    paragraph4.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph4.Range.InsertParagraphAfter();
                    Paragraph paragraph5 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph5.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range3 = paragraph5.Range;
                    range3.Text = "Billing Cycle ID\tBilling Cycle Name\tBilling Cycle Data\n";
                    List_of_dictionaries.c.CommandText = "select calendarID,calendarName from a4m.tcontractcalendarreference where branch=" + Frm_1.bank_num + " order by 1";
                    OracleDataReader oracleDataReader3 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader3.Read())
                    {
                        OracleCommand command = Frm_1.dbcon.CreateCommand();
                        command.CommandText = "select sum (statementdate1) sDate,(case WHEN sum (statementdate2)>sum (statementdate1) THEN 'Forward' WHEN (sum (statementdate2)<sum (statementdate1) and ABS(sum (statementdate2)-sum (statementdate1)) > 4) then 'Forward' ELSE 'Backward' END ) If_Holiday,CALENDARNAME,type from(select day statementdate1,null statementdate2,CALENDARNAME,type from(select substr(to_char(t.day,'dd/mm/yyyy'),0,2) day, count(substr(to_char(t.day,'dd/mm/yyyy'),0,2)) CC,t2.CALENDARNAME,'Statement_Date' Type from a4m.tcontractcalendar t,a4m.tcontractcalendarreference t2 where t.branch=" + Frm_1.bank_num + "and t.BRANCH=t2.BRANCH and t.CALENDARID=t2.CALENDARID and t.CALENDARID=" + oracleDataReader3[0].ToString() + "and t.SD=1 and to_char(t.day,'YYYYMM')>=to_char(sysdate,'YYYYMM') and to_char(t.day,'MM') <> '02' group by substr(to_char(t.day,'dd/mm/yyyy'),0,2),t2.CALENDARNAME order by CC Desc)where rownum =1 union all select null statementdate1,day statementdate2,CALENDARNAME,type from(select substr(to_char(t.day,'dd/mm/yyyy'),0,2) day, count(substr(to_char(t.day,'dd/mm/yyyy'),0,2)) CC,t2.CALENDARNAME,'Statement_Date' Type from a4m.tcontractcalendar t,a4m.tcontractcalendarreference t2 where t.branch=" + Frm_1.bank_num + "and t.BRANCH=t2.BRANCH and t.CALENDARID=t2.CALENDARID and t.CALENDARID=" + oracleDataReader3[0].ToString() + "and t.SD=1 and to_char(t.day,'YYYYMM')>=to_char(sysdate,'YYYYMM') and to_char(t.day,'MM') <> '02' group by substr(to_char(t.day,'dd/mm/yyyy'),0,2),t2.CALENDARNAME order by CC Asc)where rownum =1 )group by CALENDARNAME,type union all select sum (statementdate1) sDate,(case WHEN sum (statementdate2)>sum (statementdate1) THEN 'Forward' WHEN (sum (statementdate2)<sum (statementdate1)  and ABS(sum (statementdate2)-sum (statementdate1)) > 4) then 'Forward' ELSE 'Backward' END ) If_Holiday,CALENDARNAME,type from(select day statementdate1,null statementdate2,CALENDARNAME,type from(select substr(to_char(t.day,'dd/mm/yyyy'),0,2) day, count(substr(to_char(t.day,'dd/mm/yyyy'),0,2)) CC,t2.CALENDARNAME,'Printed_Due_Date' Type from a4m.tcontractcalendar t,a4m.tcontractcalendarreference t2 where t.branch=" + Frm_1.bank_num + "and t.BRANCH=t2.BRANCH and t.CALENDARID=t2.CALENDARID and t.CALENDARID=" + oracleDataReader3[0].ToString() + "and t.PDD=1 and to_char(t.day,'YYYYMM')>=to_char(sysdate,'YYYYMM') and to_char(t.day,'MM') <> '02' group by substr(to_char(t.day,'dd/mm/yyyy'),0,2),t2.CALENDARNAME order by CC Desc)where rownum =1 union all select null statementdate1,day statementdate2,CALENDARNAME,type from(select substr(to_char(t.day,'dd/mm/yyyy'),0,2) day, count(substr(to_char(t.day,'dd/mm/yyyy'),0,2)) CC,t2.CALENDARNAME,'Printed_Due_Date' Type from a4m.tcontractcalendar t,a4m.tcontractcalendarreference t2 where t.branch=" + Frm_1.bank_num + "and t.BRANCH=t2.BRANCH and t.CALENDARID=t2.CALENDARID and t.CALENDARID=" + oracleDataReader3[0].ToString() + "and t.PDD=1 and to_char(t.day,'YYYYMM')>=to_char(sysdate,'YYYYMM') and to_char(t.day,'MM') <> '02' group by substr(to_char(t.day,'dd/mm/yyyy'),0,2),t2.CALENDARNAME order by CC Asc)where rownum =1 )group by CALENDARNAME,type union all select sum (statementdate1) sDate,(case WHEN sum (statementdate2)>sum (statementdate1) THEN 'Forward' WHEN (sum (statementdate2)<sum (statementdate1)  and ABS(sum (statementdate2)-sum (statementdate1)) > 4) then 'Forward' ELSE 'Backward' END ) If_Holiday,CALENDARNAME,type from(select day statementdate1,null statementdate2,CALENDARNAME,type from(select substr(to_char(t.day,'dd/mm/yyyy'),0,2) day, count(substr(to_char(t.day,'dd/mm/yyyy'),0,2)) CC,t2.CALENDARNAME,'Real_Due_Date' Type from a4m.tcontractcalendar t,a4m.tcontractcalendarreference t2 where t.branch=" + Frm_1.bank_num + "and t.BRANCH=t2.BRANCH and t.CALENDARID=t2.CALENDARID and t.CALENDARID=" + oracleDataReader3[0].ToString() + " and t.RDD=1 and to_char(t.day,'YYYYMM')>=to_char(sysdate,'YYYYMM') and to_char(t.day,'MM') <> '02' group by substr(to_char(t.day,'dd/mm/yyyy'),0,2),t2.CALENDARNAME order by CC Desc)where rownum =1 union all select null statementdate1,day statementdate2,CALENDARNAME,type from(select substr(to_char(t.day,'dd/mm/yyyy'),0,2) day, count(substr(to_char(t.day,'dd/mm/yyyy'),0,2)) CC,t2.CALENDARNAME,'Real_Due_Date' Type from a4m.tcontractcalendar t,a4m.tcontractcalendarreference t2 where t.branch=" + Frm_1.bank_num + "and t.BRANCH=t2.BRANCH and t.CALENDARID=t2.CALENDARID and t.CALENDARID=" + oracleDataReader3[0].ToString() + "and t.RDD=1 and to_char(t.day,'YYYYMM')>=to_char(sysdate,'YYYYMM') and to_char(t.day,'MM') <> '02' group by substr(to_char(t.day,'dd/mm/yyyy'),0,2),t2.CALENDARNAME order by CC Asc)where rownum =1 )group by CALENDARNAME,type union all select sum (statementdate1) sDate,(case WHEN sum (statementdate2)>sum (statementdate1) THEN 'Forward' WHEN (sum (statementdate2)<sum (statementdate1)  and ABS(sum (statementdate2)-sum (statementdate1)) > 4) then 'Forward' ELSE 'Backward' END ) If_Holiday,CALENDARNAME,type from(select day statementdate1,null statementdate2,CALENDARNAME,type from(select substr(to_char(t.day,'dd/mm/yyyy'),0,2) day, count(substr(to_char(t.day,'dd/mm/yyyy'),0,2)) CC,t2.CALENDARNAME,'DAF_Date' Type from a4m.tcontractcalendar t,a4m.tcontractcalendarreference t2 where t.branch=" + Frm_1.bank_num + "and t.BRANCH=t2.BRANCH and t.CALENDARID=t2.CALENDARID and t.CALENDARID=" + oracleDataReader3[0].ToString() + "and t.DAFD=1 and to_char(t.day,'YYYYMM')>=to_char(sysdate,'YYYYMM') and to_char(t.day,'MM') <> '02' group by substr(to_char(t.day,'dd/mm/yyyy'),0,2),t2.CALENDARNAME order by CC Desc)where rownum =1 union all select null statementdate1,day statementdate2,CALENDARNAME,Type from(select substr(to_char(t.day,'dd/mm/yyyy'),0,2) day, count(substr(to_char(t.day,'dd/mm/yyyy'),0,2)) CC,t2.CALENDARNAME,'DAF_Date' Type from a4m.tcontractcalendar t,a4m.tcontractcalendarreference t2 where t.branch=" + Frm_1.bank_num + "and t.BRANCH=t2.BRANCH and t.CALENDARID=t2.CALENDARID and t.CALENDARID=" + oracleDataReader3[0].ToString() + "and t.DAFD=1 and to_char(t.day,'YYYYMM')>=to_char(sysdate,'YYYYMM') and to_char(t.day,'MM') <> '02' group by substr(to_char(t.day,'dd/mm/yyyy'),0,2),t2.CALENDARNAME order by CC Asc)where rownum =1 )group by CALENDARNAME,type";
                        OracleDataReader oracleDataReader2 = command.ExecuteReader();
                        string str1 = "";
                        while (oracleDataReader2.Read())
                            str1 = str1 + oracleDataReader2[3].ToString() + ":" + oracleDataReader2[0].ToString() + "(" + oracleDataReader2[1].ToString() + ")                                   ";
                        List_of_dictionaries.col_width = 165;
                        oracleDataReader2.Close();
                        Microsoft.Office.Interop.Word.Range range2 = range3;
                        string str2 = range2.Text + oracleDataReader3[0] + "\t" + oracleDataReader3[1] + "\t" + str1 + "\n";
                        range2.Text = str2;
                    }
                    oracleDataReader3.Close();
                    range3.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                    range3.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range3.Tables[1].Borders.Enable = 1;
                    oPara1 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                }
                if (flag7)
                {
                    oPara1.Range.Text = "\fSection 7 : Interest Settings - Credit Products Only\n";
                    oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara1.Range.Font.Bold = 0;
                    oPara1.Range.Font.Size = 9f;
                    oPara1.Format.SpaceAfter = 0.0f;
                    oPara1.Range.InsertParagraphAfter();
                    Paragraph paragraph2 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph2.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range1 = paragraph2.Range;
                    range1.Text = "Interest Name\tRate First Year\tRate After Year\n";

                    List_of_dictionaries.c.CommandText = "SELECT n name, SUM (FIRST) FIRST, NVL (SUM (second), SUM (FIRST)) second FROM (SELECT t1.NAME n, t2.PERCENTVALUE FIRST, NULL second FROM a4m.tpercentname t1, a4m.tpercentvalue t2, a4m.tpercenthistory t3 WHERE t1.BRANCH = t3.BRANCH AND t1.ID = t3.ID AND t3.CODE = t2.CODE AND t1.BRANCH = " + Frm_1.bank_num + " AND t2.COLUMNVALUE = 0 AND t3.PRCDATE IN (SELECT MAX (PRCDATE) FROM a4m.tpercenthistory h WHERE h.ID = t1.ID AND h.branch = t1.BRANCH) UNION ALL SELECT t1.NAME n, NULL FIRST, t2.PERCENTVALUE second FROM a4m.tpercentname t1, a4m.tpercentvalue t2, a4m.tpercenthistory t3 WHERE t1.BRANCH = t3.BRANCH AND t1.ID = t3.ID AND t3.CODE = t2.CODE AND t1.BRANCH = " + Frm_1.bank_num + " AND t2.COLUMNVALUE = 12 AND t3.PRCDATE IN (SELECT MAX (PRCDATE) FROM a4m.tpercenthistory h WHERE h.ID = t1.ID AND h.branch = t1.BRANCH)) GROUP BY n";

                    OracleDataReader oracleDataReader1 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader1.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range1;
                        string str = range2.Text + oracleDataReader1[0] + "\t" + oracleDataReader1[1] + "\t" + oracleDataReader1[2] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader1.Close();
                    range1.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range1.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range1.Tables[1].Borders.Enable = 1;
                    Paragraph paragraph3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph3.Range.Text = "\n--Interest Setting By Calculation Profile And Operations Groups\n";
                    paragraph3.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph3.Range.InsertParagraphAfter();
                    if (this.checkBox1.Checked)
                    {
                        Paragraph paragraph4 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        paragraph4.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                        Microsoft.Office.Interop.Word.Range range2 = paragraph4.Range;
                        range2.Text = "Profile Name\tParameter Name\tParameter Value\n";
                        List_of_dictionaries.c.CommandText = "SELECT DISTINCT tb1.branch,tb3.profilename,decode(tb2.KEY,'REPAYORDER','REPAYMENT ORDER') Parameter,DECODE (tb2.VALUE,1, 'Transaction Date Then Group Priority',2, 'Group Priority Then Transaction Date',NULL)VALUE FROM a4m.tcontractotherobjecttype tb1,a4m.tcontractotherparameters tb2,a4m.tcontractprofile tb3 WHERE     tb1.branch = " + Frm_1.bank_num + " AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND SUBSTR (tb2.KEY, 1, 10) = 'REPAYORDER' UNION ALL SELECT DISTINCT tb1.branch,tb3.profilename,decode(tb2.KEY,'REPAYDATE','REPAYMENT DATE') Parameter,DECODE (tb2.VALUE,1,'Transaction Date',2,'Posting Date',NULL)VALUE FROM a4m.tcontractotherobjecttype tb1,a4m.tcontractotherparameters tb2,a4m.tcontractprofile tb3 WHERE     tb1.branch = " + Frm_1.bank_num + " AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND SUBSTR (tb2.KEY, 1, 9) = 'REPAYDATE' UNION ALL SELECT DISTINCT tb1.branch,tb3.profilename,decode(tb2.KEY,'REPAYSETTINGS','REPAYMENT SETTINGS') Parameter,DECODE (tb2.VALUE,1, 'NOT USED',2, 'At First Pay Billed Transactions Only',3, 'Pay Transactions By Cycles',NULL)VALUE FROM a4m.tcontractotherobjecttype tb1,a4m.tcontractotherparameters tb2,a4m.tcontractprofile tb3 WHERE     tb1.branch = " + Frm_1.bank_num + " AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND SUBSTR (tb2.KEY, 1, 13) = 'REPAYSETTINGS'  order by 2,3";
                        OracleDataReader oracleDataReader2 = List_of_dictionaries.c.ExecuteReader();
                        string str1 = "";
                        while (oracleDataReader2.Read())
                        {
                            if (oracleDataReader2[1].ToString() == str1)
                            {
                                Microsoft.Office.Interop.Word.Range range3 = range2;
                                string str2 = range3.Text + " \t" + oracleDataReader2[2] + "\t" + oracleDataReader2[3] + "\n";
                                range3.Text = str2;
                            }
                            else
                            {
                                str1 = oracleDataReader2[1].ToString();
                                Microsoft.Office.Interop.Word.Range range3 = range2;
                                string str2 = range3.Text + oracleDataReader2[1] + "\t" + oracleDataReader2[2] + "\t" + oracleDataReader2[3] + "\n";
                                range3.Text = str2;
                            }
                        }
                        oracleDataReader2.Close();
                        range2.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                        range2.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                        range2.Tables[1].Borders.Enable = 1;
                        List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt).Range.InsertParagraphAfter();
                    }
                    Paragraph paragraph5 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph5.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range4 = paragraph5.Range;
                    range4.Text = "Profile Name\tGroup Name\tGroup Priority\tParameter Name\tParameter Value\n";
                    if (!this.checkBox1.Checked)
                    {
                        if (Frm_1.bank_num == "47")
                        {
                            List_of_dictionaries.c.CommandText = "SELECT profilename, groupname, priority, DECODE (UPPER (SUBSTR (KEY, 2, 11)), '_STARTDATE', 'Interest Calculation From', '_DAYSINYEAR', 'Days In Year', '_CRDPRCHIST', 'Regular Interest Rate', '_REDBASEDON', 'Use Grace Period', '_REDPRCHIST', 'Reduced Interest Rate', '_CHARGEONB', 'Charge Interest On', '_PROPRCHIST', 'Promotional Interest Rate', '_PROCALID', 'Promotional Calendar ID', '_PREPRCHIST', 'Preferential Interest Rate', '_PRECALID', 'Preferential Calendar ID', '_PREUSE', 'Usage Cycles') interest_name, VALUE interest_value FROM (SELECT DISTINCT tb1.branch, tb3.profilename, tb5.groupname, tb4.priority, tb2.KEY, CASE WHEN tb5.groupname = 'Cash' THEN 'Transaction Date' ELSE 'Payment Due Date + 1' END AS VALUE FROM a4m.tcontractotherobjecttype tb1, a4m.tcontractotherparameters tb2, a4m.tcontractprofile tb3, a4m.tcontractprofilegroup tb4, a4m.tcontractentrygroup tb5 WHERE tb1.branch = 47 AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.KEY, 2, 10) = '_STARTDATE' UNION ALL SELECT DISTINCT tb1.branch, tb3.profilename, tb5.groupname, tb4.priority, tb2.KEY, tb2.VALUE VALUE FROM a4m.tcontractotherobjecttype tb1, a4m.tcontractotherparameters tb2, a4m.tcontractprofile tb3, a4m.tcontractprofilegroup tb4, a4m.tcontractentrygroup tb5 WHERE tb1.branch = 47 AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.KEY, 2, 11) = '_DAYSINYEAR' UNION ALL SELECT DISTINCT tb1.branch, tb3.profilename, tb5.groupname, tb4.priority, tb2.KEY, DECODE (tb2.VALUE, 1, 'Full balance', 2, 'Remaining Balance', NULL) VALUE FROM a4m.tcontractotherobjecttype tb1, a4m.tcontractotherparameters tb2, a4m.tcontractprofile tb3, a4m.tcontractprofilegroup tb4, a4m.tcontractentrygroup tb5 WHERE tb1.branch = 47 AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.KEY, 2, 10) = '_CHARGEONB' UNION ALL SELECT DISTINCT tb2.branch, tb3.profilename, tb5.groupname, tb4.priority, tb2.objectcat, tn.NAME VALUE FROM a4m.tpercentbelong tb2 JOIN a4m.tpercentname tn ON tb2.branch = tn.branch AND tb2.ID = tn.ID JOIN a4m.tcontractprofile tb3 ON tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid JOIN a4m.tcontractprofilegroup tb4 ON tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid JOIN a4m.tcontractentrygroup tb5 ON tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid WHERE tb2.branch = 47 AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.objectcat, 2, 11) = '_CrdPrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID UNION ALL SELECT DISTINCT tb3.branch, tb3.profilename, tb5.groupname, tb4.priority, ' _REDBASEDON' KEY, DECODE (tb6.terms, 1, 'Min Payment Should Be Paid Till Due Date', 2, 'SD Should Be Paid Till Due Date', 3, 'SD balance is less than ' || tb6.amount) VALUE FROM a4m.tcontractotherparameters tb22, a4m.tcontractprofile tb3, a4m.tcontractprofilegroup tb4, a4m.tcontractentrygroup tb5, a4m.tcontractredintsettings tb6 WHERE tb3.branch = 47 AND tb22.branch = tb3.branch AND tb22.objectno = TO_CHAR (tb3.profileid) AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb22.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb22.KEY, 2, 11) = '_USEREDRATE' AND tb22.VALUE = 1 AND tb6.branch = tb3.branch AND tb6.profileid = tb3.profileid AND tb6.branch = tb5.branch AND TB6.GROUPID = tb5.groupid UNION ALL SELECT DISTINCT tb3.branch, tb3.profilename, tb5.groupname, tb4.priority, ' _RedPrcHist' Key, NVL (tn.NAME, 'Do Not Charge') VALUE FROM a4m.tpercentname tn, a4m.tcontractprofile tb3, a4m.tcontractprofilegroup tb4, a4m.tcontractentrygroup tb5, a4m.tcontractredintsettings tb6 WHERE tb3.branch = 47 AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND tb6.branch = tn.branch(+) AND tb6.RATEID = tn.ID(+) AND tb6.branch = tb3.branch AND tb6.profileid = tb3.profileid AND tb6.branch = tb5.branch AND TB6.GROUPID = tb5.groupid UNION ALL SELECT DISTINCT tb2.branch, tb3.profilename, tb5.groupname, tb4.priority, tb2.objectcat, tn.NAME VALUE FROM a4m.tpercentbelong tb2, a4m.tpercentname tn, a4m.tcontractprofile tb3, a4m.tcontractprofilegroup tb4, a4m.tcontractentrygroup tb5 WHERE tb2.branch = 47 AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.objectcat, 2, 11) = '_ProPrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID UNION ALL SELECT DISTINCT tb2.branch, tb3.profilename, tb5.groupname, tb4.priority, tb22.KEY, tb22.VALUE VALUE FROM a4m.tpercentbelong tb2, a4m.tpercentname tn, a4m.tcontractprofile tb3, a4m.tcontractprofilegroup tb4, a4m.tcontractentrygroup tb5, a4m.tcontractotherparameters tb22, a4m.tcontractotherobjecttype tb1 WHERE tb2.branch = 47 AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.objectcat, 2, 11) = '_ProPrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID AND tb2.ID <> -1 AND tb1.branch = tb22.branch AND tb1.objecttype = tb22.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb22.branch = tb3.branch AND tb22.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb22.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb22.KEY, 2, 11) = '_PROCALID' UNION ALL SELECT DISTINCT tb2.branch, tb3.profilename, tb5.groupname, tb4.priority, tb2.objectcat, tn.NAME VALUE FROM a4m.tpercentbelong tb2, a4m.tpercentname tn, a4m.tcontractprofile tb3, a4m.tcontractprofilegroup tb4, a4m.tcontractentrygroup tb5 WHERE tb2.branch = 47 AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.objectcat, 2, 11) = '_PrePrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID UNION ALL SELECT DISTINCT tb2.branch, tb3.profilename, tb5.groupname, tb4.priority, tb22.KEY, tb22.VALUE VALUE FROM a4m.tpercentbelong tb2, a4m.tpercentname tn, a4m.tcontractprofile tb3, a4m.tcontractprofilegroup tb4, a4m.tcontractentrygroup tb5, a4m.tcontractotherparameters tb22, a4m.tcontractotherobjecttype tb1 WHERE tb2.branch = 47 AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.objectcat, 2, 11) = '_PrePrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID AND tb2.ID <> -1 AND tb1.branch = tb22.branch AND tb1.objecttype = tb22.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb22.branch = tb3.branch AND tb22.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb22.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb22.KEY, 2, 11) = '_PRECALID' UNION ALL SELECT DISTINCT tb2.branch, tb3.profilename, tb5.groupname, tb4.priority, tb22.KEY, tb22.VALUE VALUE FROM a4m.tpercentbelong tb2, a4m.tpercentname tn, a4m.tcontractprofile tb3, a4m.tcontractprofilegroup tb4, a4m.tcontractentrygroup tb5, a4m.tcontractotherparameters tb22, a4m.tcontractotherobjecttype tb1 WHERE tb2.branch = 47 AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.objectcat, 2, 11) = '_PrePrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID AND tb2.ID <> -1 AND tb1.branch = tb22.branch AND tb1.objecttype = tb22.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb22.branch = tb3.branch AND tb22.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb22.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb22.KEY, 2, 11) = '_PREUSE') ORDER BY 1, 3, 4"; //FETCH FIRST 3000 ROWS ONLY  (for testing)
                        }
                        else
                        {
                            List_of_dictionaries.c.CommandText = "SELECT profilename,groupname,priority,DECODE (UPPER (SUBSTR (KEY, 2, 11)),'_STARTDATE', 'Interest Calculation From','_DAYSINYEAR', 'Days In Year','_CRDPRCHIST', 'Regular Interest Rate','_REDBASEDON', 'Use Grace Period','_REDPRCHIST', 'Reduced Interest Rate','_CHARGEONB', 'Charge Interest On','_PROPRCHIST', 'Promotional Interest Rate','_PROCALID', 'Promotional Calendar ID','_PREPRCHIST', 'Preferential Interest Rate','_PRECALID', 'Preferential Calendar ID','_PREUSE', 'Usage Cycles') interest_name,VALUE interest_value FROM (SELECT DISTINCT tb1.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.KEY,DECODE (tb2.VALUE,1, 'Transaction Date',2, 'Posting date',3, 'Statement date',4, 'Payment Due Date',5, 'Statement date + 15',NULL)VALUE FROM a4m.tcontractotherobjecttype tb1,a4m.tcontractotherparameters tb2,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb1.branch =" + Frm_1.bank_num + " AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.KEY, 2, 10) = '_STARTDATE' UNION ALL SELECT DISTINCT tb1.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.KEY,tb2.VALUE VALUE FROM a4m.tcontractotherobjecttype tb1,a4m.tcontractotherparameters tb2,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb1.branch =" + Frm_1.bank_num + " AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.KEY, 2, 11) = '_DAYSINYEAR' UNION ALL SELECT DISTINCT tb1.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.KEY,DECODE (tb2.VALUE,1, 'Full balance',2, 'Remaining Balance',NULL)VALUE FROM a4m.tcontractotherobjecttype tb1,a4m.tcontractotherparameters tb2,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb1.branch ='" + Frm_1.bank_num + "' AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.KEY, 2, 10) = '_CHARGEONB' UNION ALL SELECT DISTINCT tb2.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.objectcat,tn.NAME VALUE FROM a4m.tpercentbelong tb2,a4m.tpercentname tn,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb2.branch ='" + Frm_1.bank_num + "' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.objectcat, 2, 11) = '_CrdPrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID UNION ALL SELECT DISTINCT tb3.branch,tb3.profilename,tb5.groupname,tb4.priority,' _REDBASEDON' KEY,DECODE (tb6.terms,1, 'Min Payment Should Be Paid Till Due Date',2, 'SD Should Be Paid Till Due Date',3,'SD balance is less than '||tb6.amount) VALUE FROM  a4m.tcontractotherparameters tb22,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5,a4m.tcontractredintsettings tb6 WHERE     tb3.branch ='" + Frm_1.bank_num + "' AND tb22.branch = tb3.branch AND tb22.objectno = to_char(tb3.profileid) AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb22.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb22.KEY, 2, 11) = '_USEREDRATE' AND tb22.VALUE = 1 AND tb6.branch = tb3.branch AND tb6.profileid = tb3.profileid AND tb6.branch = tb5.branch AND TB6.GROUPID = tb5.groupid UNION ALL SELECT DISTINCT tb3.branch,tb3.profilename,tb5.groupname,tb4.priority,' _RedPrcHist' Key,NVL(tn.NAME,'Do Not Charge') VALUE FROM a4m.tpercentname tn,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5,a4m.tcontractredintsettings tb6 WHERE     tb3.branch = '" + Frm_1.bank_num + "' AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND tb6.branch = tn.branch(+) AND tb6.RATEID = tn.ID(+) AND tb6.branch = tb3.branch AND tb6.profileid = tb3.profileid AND tb6.branch = tb5.branch AND TB6.GROUPID = tb5.groupid UNION ALL SELECT DISTINCT tb2.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.objectcat,tn.NAME VALUE FROM a4m.tpercentbelong tb2,a4m.tpercentname tn,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb2.branch ='" + Frm_1.bank_num + "' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.objectcat, 2, 11) = '_ProPrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID UNION ALL SELECT DISTINCT tb2.branch,tb3.profilename,tb5.groupname,tb4.priority,tb22.KEY,tb22.VALUE VALUE FROM a4m.tpercentbelong tb2,a4m.tpercentname tn,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5,a4m.tcontractotherparameters tb22,a4m.tcontractotherobjecttype tb1 WHERE     tb2.branch = '" + Frm_1.bank_num + "' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.objectcat, 2, 11) = '_ProPrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID AND tb2.ID <> -1 AND tb1.branch = tb22.branch AND tb1.objecttype = tb22.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb22.branch = tb3.branch AND tb22.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb22.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb22.KEY, 2, 11) = '_PROCALID' UNION ALL SELECT DISTINCT tb2.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.objectcat,tn.NAME VALUE FROM a4m.tpercentbelong tb2,a4m.tpercentname tn,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb2.branch ='" + Frm_1.bank_num + "' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.objectcat, 2, 11) = '_PrePrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID UNION ALL SELECT DISTINCT tb2.branch,tb3.profilename,tb5.groupname,tb4.priority,tb22.KEY,tb22.VALUE VALUE FROM a4m.tpercentbelong tb2,a4m.tpercentname tn,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5,a4m.tcontractotherparameters tb22, a4m.tcontractotherobjecttype tb1 WHERE     tb2.branch = '" + Frm_1.bank_num + "' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.objectcat, 2, 11) = '_PrePrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID AND tb2.ID <> -1 AND tb1.branch = tb22.branch AND tb1.objecttype = tb22.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb22.branch = tb3.branch AND tb22.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid  AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb22.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb22.KEY, 2, 11) = '_PRECALID' UNION ALL SELECT DISTINCT tb2.branch,tb3.profilename,tb5.groupname,tb4.priority,tb22.KEY,tb22.VALUE VALUE FROM a4m.tpercentbelong tb2,a4m.tpercentname tn,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5,a4m.tcontractotherparameters tb22,a4m.tcontractotherobjecttype tb1 WHERE     tb2.branch ='" + Frm_1.bank_num + "' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.objectcat, 2, 11) = '_PrePrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID AND tb2.ID <> -1 AND tb1.branch = tb22.branch AND tb1.objecttype = tb22.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb22.branch = tb3.branch AND tb22.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb22.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb22.KEY, 2, 11) = '_PREUSE' )ORDER BY 1, 3, 4 "; //FETCH FIRST 3000 ROWS ONLY  (for testing)
                        }
                    }
                    else if (this.checkBox1.Checked)
                        List_of_dictionaries.c.CommandText = "SELECT profilename,groupname,priority,DECODE (UPPER (SUBSTR (KEY, 2, 11)),'_STARTDATE', 'Interest Calculation From','_DAYSINYEAR', 'Days In Year','_CRDPRCHIST', 'Regular Interest Rate','_REDBASEDON', 'Use Grace Period','_REDPRCHIST', 'Reduced Interest Rate','_CHARGEONB', 'Charge Interest On','_PROPRCHIST', 'Promotional Interest Rate','_PROCALID', 'Promotional Calendar ID','_PREPRCHIST', 'Preferential Interest Rate','_PRECALID', 'Preferential Calendar ID','_PREUSE', 'Usage Cycles','_CHARGEPAID','Do not Charge interest on paid amount till due date','_ONLYFIRST','In First Billing Cycle Only','_TILLCURSD','Charge Interest Till Charging Date','_UNBILLINC','Charge Interest On Unbilled Transactions Too') interest_name,VALUE interest_value FROM (SELECT DISTINCT tb1.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.KEY,DECODE (tb2.VALUE,1, 'Transaction Date',2, 'Posting date',3, 'Statement date',4, 'Payment Due Date',5, 'Statement date + 15',NULL)VALUE FROM a4m.tcontractotherobjecttype tb1,a4m.tcontractotherparameters tb2,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb1.branch =" + Frm_1.bank_num + " AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.KEY, 2, 10) = '_STARTDATE' UNION ALL SELECT DISTINCT tb1.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.KEY,tb2.VALUE VALUE FROM a4m.tcontractotherobjecttype tb1,a4m.tcontractotherparameters tb2,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb1.branch =" + Frm_1.bank_num + " AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.KEY, 2, 11) = '_DAYSINYEAR' UNION ALL SELECT DISTINCT tb1.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.KEY,DECODE (tb2.VALUE,1,'Yes','NO')VALUE FROM a4m.tcontractotherobjecttype tb1,a4m.tcontractotherparameters tb2,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb1.branch =" + Frm_1.bank_num + " AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.KEY, 2, 11) = '_CHARGEPAID'union all SELECT DISTINCT tb1.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.KEY,DECODE (tb2.VALUE,1, 'YES','NO')VALUE FROM a4m.tcontractotherobjecttype tb1,a4m.tcontractotherparameters tb2,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb1.branch =" + Frm_1.bank_num + " AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.KEY, 2, 10) = '_ONLYFIRST' union all SELECT DISTINCT tb1.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.KEY,DECODE (tb2.VALUE,1,'YES',0,'NO',NULL)VALUE FROM a4m.tcontractotherobjecttype tb1,a4m.tcontractotherparameters tb2,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb1.branch =" + Frm_1.bank_num + " AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.KEY, 2, 10) = '_TILLCURSD' union all SELECT DISTINCT tb1.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.KEY,DECODE (tb2.VALUE,1,'YES',0,'NO',NULL)VALUE FROM a4m.tcontractotherobjecttype tb1,a4m.tcontractotherparameters tb2,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb1.branch =" + Frm_1.bank_num + "AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.KEY, 1, 1) = TO_CHAR (tb5.groupid)AND SUBSTR (tb2.KEY, 2, 10) = '_UNBILLINC'union ALL SELECT DISTINCT tb1.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.KEY,DECODE (tb2.VALUE,1, 'Full balance',2, 'Remaining Balance',NULL) VALUE FROM a4m.tcontractotherobjecttype tb1,a4m.tcontractotherparameters tb2,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb1.branch =" + Frm_1.bank_num + " AND tb1.branch = tb2.branch AND tb1.objecttype = tb2.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.KEY, 2, 10) = '_CHARGEONB' UNION ALL SELECT DISTINCT tb2.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.objectcat,tn.NAME VALUE FROM a4m.tpercentbelong tb2,a4m.tpercentname tn,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb2.branch =" + Frm_1.bank_num + "AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.objectcat, 2, 11) = '_CrdPrcHist'AND tb2.branch = tn.branch AND tb2.ID = tn.ID UNION ALL SELECT DISTINCT tb3.branch,tb3.profilename,tb5.groupname,tb4.priority,' _REDBASEDON' KEY,DECODE (tb6.terms,1, 'Min Payment Should Be Paid Till Due Date',2, 'SD Should Be Paid Till Due Date',3, 'SD balance is less than ' || tb6.amount)VALUE FROM a4m.tcontractotherparameters tb22,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5,a4m.tcontractredintsettings tb6 WHERE     tb3.branch =" + Frm_1.bank_num + " AND tb22.branch = tb3.branch AND tb22.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb22.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb22.KEY, 2, 11) = '_USEREDRATE' AND tb22.VALUE = 1 AND tb6.branch = tb3.branch AND tb6.profileid = tb3.profileid AND tb6.branch = tb5.branch AND TB6.GROUPID = tb5.groupid UNION ALL SELECT DISTINCT tb3.branch,tb3.profilename,tb5.groupname,tb4.priority,' _RedPrcHist' Key,NVL (tn.NAME, 'Do Not Charge') VALUE FROM a4m.tpercentname tn,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5,a4m.tcontractredintsettings tb6 WHERE     tb3.branch =" + Frm_1.bank_num + " AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND tb6.branch = tn.branch(+) AND tb6.RATEID = tn.ID(+) AND tb6.branch = tb3.branch AND tb6.profileid = tb3.profileid AND tb6.branch = tb5.branch AND TB6.GROUPID = tb5.groupid UNION ALL SELECT DISTINCT tb2.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.objectcat,tn.NAME VALUE FROM a4m.tpercentbelong tb2,a4m.tpercentname tn,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb2.branch =" + Frm_1.bank_num + " AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid)AND SUBSTR (tb2.objectcat, 2, 11) = '_ProPrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID UNION ALL SELECT DISTINCT tb2.branch,tb3.profilename,tb5.groupname,tb4.priority,tb22.KEY,tb22.VALUE VALUE FROM a4m.tpercentbelong tb2,a4m.tpercentname tn,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5,a4m.tcontractotherparameters tb22,a4m.tcontractotherobjecttype tb1 WHERE     tb2.branch =" + Frm_1.bank_num + " AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid)AND SUBSTR (tb2.objectcat, 2, 11) = '_ProPrcHist'AND tb2.branch = tn.branch AND tb2.ID = tn.ID AND tb2.ID <> -1 AND tb1.branch = tb22.branch AND tb1.objecttype = tb22.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb22.branch = tb3.branch AND tb22.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb22.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb22.KEY, 2, 11) = '_PROCALID' UNION ALL SELECT DISTINCT tb2.branch,tb3.profilename,tb5.groupname,tb4.priority,tb2.objectcat,tn.NAME VALUE FROM a4m.tpercentbelong tb2,a4m.tpercentname tn,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5 WHERE     tb2.branch =" + Frm_1.bank_num + " AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid)AND SUBSTR (tb2.objectcat, 2, 11) = '_PrePrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID UNION ALL SELECT DISTINCT tb2.branch,tb3.profilename,tb5.groupname,tb4.priority,tb22.KEY,tb22.VALUE VALUE FROM a4m.tpercentbelong tb2,a4m.tpercentname tn,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5,a4m.tcontractotherparameters tb22,a4m.tcontractotherobjecttype tb1 WHERE     tb2.branch =" + Frm_1.bank_num + " AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.objectcat, 2, 11) = '_PrePrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID AND tb2.ID <> -1 AND tb1.branch = tb22.branch AND tb1.objecttype = tb22.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb22.branch = tb3.branch AND tb22.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb22.KEY, 1, 1) = TO_CHAR (tb5.groupid)AND SUBSTR (tb22.KEY, 2, 11) = '_PRECALID' UNION ALL SELECT DISTINCT tb2.branch,tb3.profilename,tb5.groupname,tb4.priority,tb22.KEY,tb22.VALUE VALUE FROM a4m.tpercentbelong tb2,a4m.tpercentname tn,a4m.tcontractprofile tb3,a4m.tcontractprofilegroup tb4,a4m.tcontractentrygroup tb5,a4m.tcontractotherparameters tb22,a4m.tcontractotherobjecttype tb1 WHERE     tb2.branch =" + Frm_1.bank_num + " AND tb2.branch = tb3.branch AND tb2.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb2.objectcat, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb2.objectcat, 2, 11) = '_PrePrcHist' AND tb2.branch = tn.branch AND tb2.ID = tn.ID AND tb2.ID <> -1 AND tb1.branch = tb22.branch AND tb1.objecttype = tb22.objecttype AND tb1.objectsign LIKE 'CONTRACT_PROFILE' AND tb22.branch = tb3.branch AND tb22.objectno = tb3.profileid AND tb3.branch = tb4.branch AND tb3.profileid = tb4.profileid AND tb4.branch = tb5.branch AND tb4.groupid = tb5.groupid AND SUBSTR (tb22.KEY, 1, 1) = TO_CHAR (tb5.groupid) AND SUBSTR (tb22.KEY, 2, 11) = '_PREUSE') ORDER BY 1, 3, 4";
                    OracleDataReader oracleDataReader3 = List_of_dictionaries.c.ExecuteReader();
                    string str3 = "";
                    string str4 = "";
                    string str5 = "";
                    string og_filename = List_of_dictionaries.filename.ToString();
                    int l_count = 0;
                    int splitNum = 0; //Split threshold
                    while (oracleDataReader3.Read())
                    {
                        l_count++;
                        if (oracleDataReader3[0].ToString() == str3 && oracleDataReader3[1].ToString() == str4 && oracleDataReader3[2].ToString() == str5)
                        {
                            Microsoft.Office.Interop.Word.Range range2 = range4;
                            string str1 = range2.Text + " \t \t \t" + oracleDataReader3[3] + "\t" + oracleDataReader3[4] + "\n";
                            range2.Text = str1;
                            range2 = null;
                        }
                        else
                        {
                            str3 = oracleDataReader3[0].ToString();
                            str4 = oracleDataReader3[1].ToString();
                            str5 = oracleDataReader3[2].ToString();
                            Microsoft.Office.Interop.Word.Range range2 = range4;
                            string str1 = range2.Text + oracleDataReader3[0] + "\t" + oracleDataReader3[1] + "\t" + oracleDataReader3[2] + "\t" + oracleDataReader3[3] + "\t" + oracleDataReader3[4] + "\n";
                            range2.Text = str1;
                        }
                        if (l_count % 2000 == 0)//use direct equality for split threshold, use mod % for splitting every x entries
                        {
                            range4.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                            range4.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                            range4.Tables[1].Borders.Enable = 1;
                            string filename_part = og_filename + "_part" + ++splitNum;
                            List_of_dictionaries.filename = og_filename + "_part" + (splitNum + 1);
                            List_of_dictionaries.oDoc.SaveAs(filename_part, WdSaveFormat.wdFormatDocumentDefault, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                            List_of_dictionaries.oDoc.Close(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                            oWord = (Word.Application)new Word.Application();
                            oWord.Visible = false;
                            oWord.KeyboardLatin();
                            oWord.Keyboard(2057);
                            oDoc = (_Document)List_of_dictionaries.oWord.Documents.Add(ref m_objOpt, ref List_of_dictionaries.m_objOpt, ref m_objOpt, ref m_objOpt);
                            oDoc.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            oDoc.PageSetup.RightMargin = 20f;
                            oDoc.PageSetup.LeftMargin = 20f;
                            oDoc.PageSetup.TopMargin = 20f;
                            oDoc.PageSetup.BottomMargin = 20f;
                            oDoc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

                            oDoc.ShowGrammaticalErrors = false;
                            oDoc.ShowRevisions = false;
                            oDoc.ShowSpellingErrors = false;

                            oPara1 = oDoc.Content.Paragraphs.Add(ref m_objOpt);
                            oPara1.Range.Text = Frm_1.bank_fiid + " Settings Report On " + Frm_1.dbname + " (Cont.)";
                            oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            oPara1.Range.Font.Name = "Verdana";
                            oPara1.Range.Font.Color = WdColor.wdColorDarkBlue;
                            oPara1.Range.Font.Size = 15f;
                            oPara1.Format.SpaceAfter = 18f;
                            oPara1.Range.InsertParagraphAfter();
                            oPara1.Range.Font.Bold = 0;
                            oPara1.Range.Font.Size = 9f;
                            oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                            oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            oPara1.Range.Text = "\n--Interest Setting By Calculation Profile And Operations Groups (Cont.)\n";// "\fSection 7 : Interest Settings - Credit Products Only\n";
                            oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                            oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            oPara1.Range.Font.Bold = 0;
                            oPara1.Range.Font.Size = 9f;
                            oPara1.Format.SpaceAfter = 0.0f;
                            oPara1.Range.InsertParagraphAfter();
                            paragraph5 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                            paragraph5.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                            range4 = paragraph5.Range;
                            range4.Text = "Profile Name\tGroup Name\tGroup Priority\tParameter Name\tParameter Value\n";
                        }
                    }
                    oracleDataReader3.Close();
                    range4.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range4.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range4.Tables[1].Borders.Enable = 1;
                    paragraph12 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    Paragraph paragraph6 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph6.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range5 = paragraph6.Range;
                    range5.Text = "Contract Name\tContract Profile\tContract Profile Currency\n";
                    List_of_dictionaries.c.CommandText = "SELECT NAME, PROFILE, currency, Prcdate INSPCNTDATE FROM ( ( SELECT X.Name NAME, N.Name PROFILE, NULL currency, MAX (H.Prcdate) Prcdate FROM A4M.Tpercentbelong B, A4M.Tcontracttype X, A4M.Tpercentname N, A4M.Tpercenthistory H, A4M.Tpercentvalue V WHERE B.Branch = " + Frm_1.bank_num + " AND B.Category = 1 AND B.Branch = X.Branch AND B.Objectno = X.TYPE AND B.Branch = N.Branch AND B.Id = N.Id AND B.Id <> -1 AND B.Branch = H.Branch AND B.Id = H.Id AND H.Code = V.Code AND X.SCHEMATYPE = 3 AND SUBSTR (X.Status, 1, 1) = '1' AND ( startdate IS NULL OR startdate <= TO_DATE ('16.01.2024', 'DD.MM.YYYY')) AND ( enddate IS NULL OR enddate >= TO_DATE ('16.01.2024', 'DD.MM.YYYY')) GROUP BY X.Name, N.Name UNION ALL SELECT DISTINCT tcontracttype.name, tcontractprofile.PROFILENAME, 'Domestic' Currency, TCONTRACTTYPE.STARTDATE FROM a4m.tcontractprofile, a4m.tcontracttype, a4m.tcontracttypeparameters WHERE tcontracttype.branch = tcontracttypeparameters.branch AND tcontracttype.TYPE = tcontracttypeparameters.CONTRACTTYPE AND tcontracttypeparameters.KEY LIKE 'PROFILEDOM' AND tcontracttypeparameters.branch = tcontractprofile.branch AND (tcontracttypeparameters.VALUE) = (tcontractprofile.PROFILEID) AND tcontracttype.branch = " + Frm_1.bank_num + " AND tcontracttype.status LIKE '1' UNION ALL SELECT tcontracttype.name, tcontractprofile.PROFILENAME, 'International' Currency, TCONTRACTTYPE.STARTDATE FROM a4m.tcontractprofile, a4m.tcontracttype, a4m.tcontracttypeparameters WHERE tcontracttype.branch = tcontracttypeparameters.branch AND tcontracttype.TYPE = tcontracttypeparameters.CONTRACTTYPE AND tcontracttypeparameters.KEY LIKE 'PROFILEINT' AND tcontracttypeparameters.branch = tcontractprofile.branch AND (tcontracttypeparameters.VALUE) = (tcontractprofile.PROFILEID) AND tcontracttype.branch = " + Frm_1.bank_num + " AND tcontracttype.status LIKE '1'))";
                    OracleDataReader oracleDataReader4 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader4.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range5;
                        string str1 = range2.Text + oracleDataReader4[0] + "\t" + oracleDataReader4[1] + "\t" + oracleDataReader4[2] + "\n";
                        range2.Text = str1;
                    }
                    oracleDataReader4.Close();
                    range5.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range5.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range5.Tables[1].Borders.Enable = 1;
                    oPara1 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                }
                if (flag8)
                {
                    oPara1.Range.Text = "\fSection 8 : Allowable Overlimit + Overlimit Fees - Credit Products Only\n";
                    oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara1.Range.Font.Bold = 0;
                    oPara1.Range.Font.Size = 9f;
                    oPara1.Format.SpaceAfter = 0.0f;
                    oPara1.Range.InsertParagraphAfter();
                    Paragraph paragraph2 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph2.Range.Text = "\nS8:1-Allowable Overdue\n";
                    paragraph2.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph2.Range.InsertParagraphAfter();
                    Paragraph paragraph3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph3.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range1 = paragraph3.Range;
                    range1.Text = "Name\tFlat Amount\tPercentage\tcalculation Type\tCurrency\n";
                    List_of_dictionaries.c.CommandText = "select v0 Profile,nvl(sum(v1),0) Amount,nvl(sum(v2),0) Percentage,decode(sum(v3),1,'Not Allowable',2,'As Total Of',3,'As Maximum Between',4,'As Minimum Between') Type,c currency from (select distinct tcontractprofile.PROFILENAME v0,tcontractotherparameters.VALUE v1,null v2 ,null v3,tcontractprofile.CURRENCY c from a4m.tcontractotherparameters,a4m.tcontractprofile,a4m.tcontractotherobjecttype where tcontractotherparameters.branch=" + Frm_1.bank_num + " and Upper(tcontractotherparameters.KEY) like 'OVDVALALLOWAMOUNT' and tcontractprofile.branch=tcontractotherparameters.branch and tcontractprofile.PROFILEID=tcontractotherparameters.OBJECTNO and tcontractotherobjecttype.branch=tcontractotherparameters.branch and tcontractotherobjecttype.OBJECTTYPE=tcontractotherparameters.OBJECTTYPE union all select distinct tcontractprofile.PROFILENAME v0,null v1,tcontractotherparameters.VALUE v2 ,null v3,tcontractprofile.CURRENCY c from a4m.tcontractotherparameters,a4m.tcontractprofile,a4m.tcontractotherobjecttype where tcontractotherparameters.branch=" + Frm_1.bank_num + " and Upper(tcontractotherparameters.KEY) like 'OVDVALALLOWPRC'and tcontractprofile.branch=tcontractotherparameters.branch and tcontractprofile.PROFILEID=tcontractotherparameters.OBJECTNO and tcontractotherobjecttype.branch=tcontractotherparameters.branch and tcontractotherobjecttype.OBJECTTYPE=tcontractotherparameters.OBJECTTYPE union all select distinct tcontractprofile.PROFILENAME v0,null v1,null v2 ,tcontractotherparameters.VALUE v3,tcontractprofile.CURRENCY c from a4m.tcontractotherparameters,a4m.tcontractprofile,a4m.tcontractotherobjecttype where tcontractotherparameters.branch=" + Frm_1.bank_num + " and Upper(tcontractotherparameters.KEY) like 'OVDVALALLOW'and tcontractprofile.branch=tcontractotherparameters.branch and tcontractprofile.PROFILEID=tcontractotherparameters.OBJECTNO and tcontractotherobjecttype.branch=tcontractotherparameters.branch and tcontractotherobjecttype.OBJECTTYPE=tcontractotherparameters.OBJECTTYPE )group by v0,c order by 1";
                    OracleDataReader oracleDataReader1 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader1.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range1;
                        string str = range2.Text + oracleDataReader1[0] + "\t" + oracleDataReader1[1] + "\t" + oracleDataReader1[2] + "\t" + oracleDataReader1[3] + "\t" + oracleDataReader1[4] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader1.Close();
                    range1.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range1.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range1.Tables[1].Borders.Enable = 1;
                    Paragraph paragraph4 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph4.Range.Text = "\nS8:2-Overdue Fees\n";
                    paragraph4.Range.InsertParagraphAfter();
                    Paragraph paragraph5 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph5.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range3 = paragraph5.Range;
                    range3.Text = "Name\tFlat Amount\tmaxamount\tminamount\tPercentage\tCalculation Type\tCurrency\tOverdue Fee Based On\n";
                    List_of_dictionaries.c.CommandText = "SELECT v0 PROFILE,SUM (v1) amount,Sum(v6) maxamount,Sum(v5) minamount,SUM (v2) percentage,DECODE (SUM (v3),1, 'Not Used',2, 'Charge Fee as Total Of',3, 'Charge Fee as maximum between',4, 'Charge Fee as minimum between')TYPE,c currency,DECODE (SUM (v4),  1, 'Statement Date',  2, 'Payment Due Date') TYPE2  FROM (SELECT DISTINCT tcontractprofile.profilename v0,tcontractotherparameters.VALUE v1,NULL v2,NULL v3,NULL v4,NULL V5,NULL V6,tcontractprofile.currency c FROM a4m.tcontractotherparameters,a4m.tcontractprofile,a4m.tcontractotherobjecttype WHERE tcontractotherparameters.branch =" + Frm_1.bank_num + "AND UPPER (tcontractotherparameters.KEY) LIKE 'OVDFEEAMOUNT' AND tcontractprofile.branch = tcontractotherparameters.branch AND tcontractprofile.profileid =tcontractotherparameters.objectno AND tcontractotherobjecttype.branch = tcontractotherparameters.branch AND tcontractotherobjecttype.objecttype =tcontractotherparameters.objecttype UNION ALL SELECT DISTINCT tcontractprofile.profilename v0, NULL v1,tcontractotherparameters.VALUE v2,NULL v3,NULL v4,NULL V5,NULL V6,tcontractprofile.currency c FROM a4m.tcontractotherparameters,a4m.tcontractprofile,a4m.tcontractotherobjecttype WHERE tcontractotherparameters.branch =" + Frm_1.bank_num + "AND UPPER (tcontractotherparameters.KEY) LIKE 'OVDFEEPRC' AND tcontractprofile.branch = tcontractotherparameters.branch AND tcontractprofile.profileid = tcontractotherparameters.objectno AND tcontractotherobjecttype.branch = tcontractotherparameters.branch AND tcontractotherobjecttype.objecttype = tcontractotherparameters.objecttype UNION ALL SELECT DISTINCT tcontractprofile.profilename v0, NULL v1,NULL v2,tcontractotherparameters.VALUE v3,NULL v4,NULL V5,NULL V6,tcontractprofile.currency c FROM a4m.tcontractotherparameters, a4m.tcontractprofile,a4m.tcontractotherobjecttype WHERE tcontractotherparameters.branch = " + Frm_1.bank_num + "AND UPPER (tcontractotherparameters.KEY) LIKE 'OVDFEETYPE' AND tcontractprofile.branch = tcontractotherparameters.branch AND tcontractprofile.profileid = tcontractotherparameters.objectno AND tcontractotherobjecttype.branch = tcontractotherparameters.branch AND tcontractotherobjecttype.objecttype = tcontractotherparameters.objecttype UNION ALL SELECT DISTINCT tcontractprofile.profilename v0, NULL v1,NULL v2,NULL v3,tcontractotherparameters.VALUE v4,NULL V5,NULL V6,tcontractprofile.currency c FROM a4m.tcontractotherparameters,a4m.tcontractprofile,a4m.tcontractotherobjecttype WHERE tcontractotherparameters.branch = " + Frm_1.bank_num + " AND UPPER (tcontractotherparameters.KEY) LIKE 'OVDFEEDATE' AND tcontractprofile.branch = tcontractotherparameters.branch AND tcontractprofile.profileid =tcontractotherparameters.objectno AND tcontractotherobjecttype.branch =tcontractotherparameters.branch AND tcontractotherobjecttype.objecttype =tcontractotherparameters.objecttype UNION ALL SELECT DISTINCT tcontractprofile.profilename v0,NULL v1,NULL v2,NULL v3,NULL v4,tcontractotherparameters.VALUE V5,NULL V6,tcontractprofile.currency c FROM a4m.tcontractotherparameters,a4m.tcontractprofile,a4m.tcontractotherobjecttype WHERE tcontractotherparameters.branch = " + Frm_1.bank_num + " AND UPPER (tcontractotherparameters.KEY) LIKE 'OVDMINAMOUNT' AND tcontractprofile.branch = tcontractotherparameters.branch AND tcontractprofile.profileid = tcontractotherparameters.objectno AND tcontractotherobjecttype.branch =tcontractotherparameters.branch AND tcontractotherobjecttype.objecttype =tcontractotherparameters.objecttype UNION ALL SELECT DISTINCT tcontractprofile.profilename v0,NULL v1,NULL v2,NULL v3,NULL v4,NULL V5,tcontractotherparameters.VALUE  V6,tcontractprofile.currency c FROM a4m.tcontractotherparameters,a4m.tcontractprofile,a4m.tcontractotherobjecttype WHERE tcontractotherparameters.branch = " + Frm_1.bank_num + " AND UPPER (tcontractotherparameters.KEY) LIKE 'OVDMAXAMOUNT' AND tcontractprofile.branch = tcontractotherparameters.branch AND tcontractprofile.profileid =tcontractotherparameters.objectno AND tcontractotherobjecttype.branch = tcontractotherparameters.branch AND tcontractotherobjecttype.objecttype = tcontractotherparameters.objecttype)GROUP BY v0, c ORDER BY 1";
                    // List_of_dictionaries.c.CommandText = "SELECT   v0 PROFILE, SUM (v1) amount, SUM (v2) percentage,DECODE (SUM (v3),1, 'Not Used',2, 'Charge Fee as Total Of',3, 'Charge Fee as maximum between',4, 'Charge Fee as minimum between') TYPE,c currency,DECODE (SUM (v4),1, 'Statement Date',2, 'Payment Due Date') TYPE2 FROM (SELECT DISTINCT tcontractprofile.profilename v0,tcontractotherparameters.VALUE v1, NULL v2, NULL v3,null v4,tcontractprofile.currency c FROM a4m.tcontractotherparameters, a4m.tcontractprofile,a4m.tcontractotherobjecttype WHERE tcontractotherparameters.branch =" + Frm_1.bank_num + "AND UPPER (tcontractotherparameters.KEY) LIKE 'OVDFEEAMOUNT'AND tcontractprofile.branch =tcontractotherparameters.branch AND tcontractprofile.profileid =tcontractotherparameters.objectno AND tcontractotherobjecttype.branch =tcontractotherparameters.branch AND tcontractotherobjecttype.objecttype =tcontractotherparameters.objecttype UNION ALL SELECT DISTINCT tcontractprofile.profilename v0, NULL v1,tcontractotherparameters.VALUE v2, NULL v3,null v4,tcontractprofile.currency c FROM a4m.tcontractotherparameters, a4m.tcontractprofile,a4m.tcontractotherobjecttype WHERE tcontractotherparameters.branch =" + Frm_1.bank_num + "AND UPPER (tcontractotherparameters.KEY) LIKE'OVDFEEPRC'AND tcontractprofile.branch =tcontractotherparameters.branch AND tcontractprofile.profileid =tcontractotherparameters.objectno AND tcontractotherobjecttype.branch =tcontractotherparameters.branch AND tcontractotherobjecttype.objecttype =tcontractotherparameters.objecttype UNION ALL SELECT DISTINCT tcontractprofile.profilename v0, NULL v1, NULL v2,tcontractotherparameters.VALUE v3,null v4,tcontractprofile.currency c FROM a4m.tcontractotherparameters, a4m.tcontractprofile, a4m.tcontractotherobjecttype WHERE tcontractotherparameters.branch =" + Frm_1.bank_num + "AND UPPER (tcontractotherparameters.KEY) LIKE 'OVDFEETYPE' AND tcontractprofile.branch =tcontractotherparameters.branch AND tcontractprofile.profileid =tcontractotherparameters.objectno AND tcontractotherobjecttype.branch = tcontractotherparameters.branch AND tcontractotherobjecttype.objecttype = tcontractotherparameters.objecttype UNION ALL SELECT DISTINCT tcontractprofile.profilename v0, NULL v1,null v2, NULL v3,tcontractotherparameters.VALUE v4,tcontractprofile.currency c FROM a4m.tcontractotherparameters,a4m.tcontractprofile,a4m.tcontractotherobjecttype WHERE tcontractotherparameters.branch =" + Frm_1.bank_num + "AND UPPER (tcontractotherparameters.KEY) LIKE 'OVDFEEDATE'AND tcontractprofile.branch =tcontractotherparameters.branch AND tcontractprofile.profileid =tcontractotherparameters.objectno AND tcontractotherobjecttype.branch =tcontractotherparameters.branch AND tcontractotherobjecttype.objecttype =tcontractotherparameters.objecttype)GROUP BY v0, c ORDER BY 1";
                    //error >> Word has encountered a problem 
                    OracleDataReader oracleDataReader2 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader2.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range3;
                        string str = range2.Text + oracleDataReader2[0] + "\t" + oracleDataReader2[1] + "\t" + oracleDataReader2[2] + "\t" + oracleDataReader2[3] + "\t" + oracleDataReader2[4] + "\t" + oracleDataReader2[5] + "\t" + oracleDataReader2[6] + "\t" + oracleDataReader2[7] + "\n";

                        range2.Text = str;
                    }
                    oracleDataReader2.Close();
                    range3.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, 100, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range3.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range3.Tables[1].Borders.Enable = 1;
                }
                if (flag9)
                {
                    oPara1.Range.Text = "\fSection 9 : Allowable Overdue + Overdue Fees - Credit Products Only\n";
                    oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara1.Range.Font.Bold = 0;
                    oPara1.Range.Font.Size = 9f;
                    oPara1.Format.SpaceAfter = 0.0f;
                    oPara1.Range.InsertParagraphAfter();
                    Paragraph paragraph6 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph6.Range.Text = "\nS9:1-Allowable Overlimit\n";
                    paragraph6.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph6.Range.InsertParagraphAfter();
                    Paragraph paragraph7 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph7.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range4 = paragraph7.Range;
                    range4.Text = "Name\tFlat Amount\tPercentage\tCalculation Type\tCurrency\n";
                    List_of_dictionaries.c.CommandText = "select v0 Profile,sum(v1) Amount,sum(v2) Percentage,decode(sum(v3),1,'As Maximum Between',2,'As Minimum Between') Type,c currency from (select distinct tcontractprofile.PROFILENAME v0,tcontractotherparameters.VALUE v1,null v2 ,null v3,tcontractprofile.CURRENCY c from a4m.tcontractotherparameters,a4m.tcontractprofile,a4m.tcontractotherobjecttype where tcontractotherparameters.branch=" + Frm_1.bank_num + " and Upper(tcontractotherparameters.KEY) like 'OVRLMTALLOWAMOUNT' and tcontractprofile.branch=tcontractotherparameters.branch and tcontractprofile.PROFILEID=tcontractotherparameters.OBJECTNO and tcontractotherobjecttype.branch=tcontractotherparameters.branch and tcontractotherobjecttype.OBJECTTYPE=tcontractotherparameters.OBJECTTYPE union all select distinct tcontractprofile.PROFILENAME v0,null v1,tcontractotherparameters.VALUE v2 ,null v3,tcontractprofile.CURRENCY c from a4m.tcontractotherparameters,a4m.tcontractprofile,a4m.tcontractotherobjecttype where tcontractotherparameters.branch=" + Frm_1.bank_num + " and Upper(tcontractotherparameters.KEY) like 'OVRLMTALLOWPRC'and tcontractprofile.branch=tcontractotherparameters.branch and tcontractprofile.PROFILEID=tcontractotherparameters.OBJECTNO and tcontractotherobjecttype.branch=tcontractotherparameters.branch and tcontractotherobjecttype.OBJECTTYPE=tcontractotherparameters.OBJECTTYPE union all select distinct tcontractprofile.PROFILENAME v0,null v1,null v2 ,tcontractotherparameters.VALUE v3,tcontractprofile.CURRENCY c from a4m.tcontractotherparameters,a4m.tcontractprofile,a4m.tcontractotherobjecttype where tcontractotherparameters.branch=" + Frm_1.bank_num + " and Upper(tcontractotherparameters.KEY) like 'OVRLMTALLOW'and tcontractprofile.branch=tcontractotherparameters.branch and tcontractprofile.PROFILEID=tcontractotherparameters.OBJECTNO and tcontractotherobjecttype.branch=tcontractotherparameters.branch and tcontractotherobjecttype.OBJECTTYPE=tcontractotherparameters.OBJECTTYPE )group by v0,c order by 1";
                    OracleDataReader oracleDataReader3 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader3.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range4;
                        string str = range2.Text + oracleDataReader3[0] + "\t" + oracleDataReader3[1] + "\t" + oracleDataReader3[2] + "\t" + oracleDataReader3[3] + "\t" + oracleDataReader3[4] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader3.Close();
                    range4.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range4.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range4.Tables[1].Borders.Enable = 1;
                    Paragraph paragraph8 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph8.Range.Text = "\nS9:2-Overlimit Fee\n";
                    paragraph8.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph8.Range.InsertParagraphAfter();
                    Paragraph paragraph9 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph9.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range5 = paragraph9.Range;
                    range5.Text = "Name\tFlat Amount\tPercentage\tCalculation Type\tCurrency\tOverlimit Fee Based On\n";
                    List_of_dictionaries.c.CommandText = "SELECT   v0 PROFILE, SUM (v1) amount, SUM (v2) percentage,DECODE (SUM (v3),1, 'Not Used',3, 'Charge Fee as Total Of',4, 'Charge Fee as maximum between',5, 'Charge Fee as minimum between',sum(v3)) TYPE,c currency,DECODE (SUM (v4),1, 'Limit Excess within Cycle',2, 'Limit Excess on Statement Date') TYPE2 FROM (SELECT DISTINCT tcontractprofile.profilename v0,tcontractotherparameters.VALUE v1, NULL v2, NULL v3,null v4,tcontractprofile.currency c FROM a4m.tcontractotherparameters, a4m.tcontractprofile,a4m.tcontractotherobjecttype WHERE tcontractotherparameters.branch =" + Frm_1.bank_num + "AND UPPER (tcontractotherparameters.KEY) LIKE 'OVRLMTAMOUNT'AND tcontractprofile.branch =tcontractotherparameters.branch AND tcontractprofile.profileid =tcontractotherparameters.objectno AND tcontractotherobjecttype.branch =tcontractotherparameters.branch AND tcontractotherobjecttype.objecttype =tcontractotherparameters.objecttype UNION ALL SELECT DISTINCT tcontractprofile.profilename v0, NULL v1,tcontractotherparameters.VALUE v2, NULL v3,null v4,tcontractprofile.currency c FROM a4m.tcontractotherparameters, a4m.tcontractprofile,a4m.tcontractotherobjecttype WHERE tcontractotherparameters.branch =" + Frm_1.bank_num + "AND UPPER (tcontractotherparameters.KEY) LIKE'OVRLMTPRC'AND tcontractprofile.branch =tcontractotherparameters.branch AND tcontractprofile.profileid =tcontractotherparameters.objectno AND tcontractotherobjecttype.branch =tcontractotherparameters.branch AND tcontractotherobjecttype.objecttype =tcontractotherparameters.objecttype UNION ALL SELECT DISTINCT tcontractprofile.profilename v0, NULL v1, NULL v2,tcontractotherparameters.VALUE v3,null v4,tcontractprofile.currency c FROM a4m.tcontractotherparameters, a4m.tcontractprofile, a4m.tcontractotherobjecttype WHERE tcontractotherparameters.branch =" + Frm_1.bank_num + "AND UPPER (tcontractotherparameters.KEY) LIKE 'OVRLMTTYPE' AND tcontractprofile.branch =tcontractotherparameters.branch AND tcontractprofile.profileid =tcontractotherparameters.objectno AND tcontractotherobjecttype.branch = tcontractotherparameters.branch AND tcontractotherobjecttype.objecttype = tcontractotherparameters.objecttype UNION ALL SELECT DISTINCT tcontractprofile.profilename v0, NULL v1,null v2, NULL v3,tcontractotherparameters.VALUE v4,tcontractprofile.currency c FROM a4m.tcontractotherparameters,a4m.tcontractprofile,a4m.tcontractotherobjecttype WHERE tcontractotherparameters.branch =" + Frm_1.bank_num + "AND UPPER (tcontractotherparameters.KEY) LIKE 'OVRLMTBASED'AND tcontractprofile.branch =tcontractotherparameters.branch AND tcontractprofile.profileid =tcontractotherparameters.objectno AND tcontractotherobjecttype.branch =tcontractotherparameters.branch AND tcontractotherobjecttype.objecttype =tcontractotherparameters.objecttype)GROUP BY v0, c ORDER BY 1";
                    OracleDataReader oracleDataReader4 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader4.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range5;
                        string str = range2.Text + oracleDataReader4[0] + "\t" + oracleDataReader4[1] + "\t" + oracleDataReader4[2] + "\t" + oracleDataReader4[3] + "\t" + oracleDataReader4[4] + "\t" + oracleDataReader4[5] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader4.Close();
                    range5.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range5.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range5.Tables[1].Borders.Enable = 1;

                    paragraph12 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    Paragraph paragraph10 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph10.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range6 = paragraph10.Range;
                    range6.Text = "Contract Name\tContract Profile\tContract Profile Currency\n";
                    List_of_dictionaries.c.CommandText = "select tcontracttype.name,tcontractprofile.PROFILENAME ,'Domestic' Currency from a4m.tcontractprofile,a4m.tcontracttype,a4m.tcontracttypeparameters where tcontracttype.branch=tcontracttypeparameters.branch and tcontracttype.TYPE=tcontracttypeparameters.CONTRACTTYPE and tcontracttypeparameters.KEY like 'PROFILEDOM'and tcontracttypeparameters.branch=tcontractprofile.branch and tcontracttypeparameters.VALUE=tcontractprofile.PROFILEID and tcontracttype.branch=" + Frm_1.bank_num + " and tcontracttype.status like '1' union all select tcontracttype.name,tcontractprofile.PROFILENAME ,'International' Currency from a4m.tcontractprofile,a4m.tcontracttype,a4m.tcontracttypeparameters where tcontracttype.branch=tcontracttypeparameters.branch and tcontracttype.TYPE=tcontracttypeparameters.CONTRACTTYPE and tcontracttypeparameters.KEY like 'PROFILEINT' and tcontracttypeparameters.branch=tcontractprofile.branch and tcontracttypeparameters.VALUE=tcontractprofile.PROFILEID and tcontracttype.branch=" + Frm_1.bank_num + " and tcontracttype.status like '1' order by 1,3";
                    OracleDataReader oracleDataReader5 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader5.Read())
                    {

                        Microsoft.Office.Interop.Word.Range range2 = range6;
                        string str = range2.Text + oracleDataReader5[0] + "\t" + oracleDataReader5[1] + "\t" + oracleDataReader5[2] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader5.Close();
                    range6.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range6.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range6.Tables[1].Borders.Enable = 1;
                }
                if (flag10)
                {
                    oPara1.Range.Text = "\fSection 10 : Credit Shield - Credit Products Only\n";
                    oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara1.Range.Font.Bold = 0;
                    oPara1.Range.Font.Size = 9f;
                    oPara1.Format.SpaceAfter = 0.0f;
                    oPara1.Range.InsertParagraphAfter();
                    Paragraph paragraph11 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph11.Range.Text = "\nS10:Credit Shield - Credit Products Only\n";
                    paragraph11.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph11.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    paragraph11.Range.Font.Bold = 0;
                    paragraph11.Range.Font.Size = 9f;
                    paragraph11.Format.SpaceAfter = 6f;
                    paragraph11.Range.InsertParagraphAfter();

                    //Islam Atta 
                    List_of_dictionaries.c.CommandText = "select type,name,decode(status,'1','Active','Inactive') Status from a4m.tcontracttype where branch=" + Frm_1.bank_num + " order by 1";
                    OracleDataReader oracleDataReader03 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader03.Read())
                    {
                        Paragraph paragraph13 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        paragraph13.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                        paragraph13.Range.SetRange(0, 20000);
                        Microsoft.Office.Interop.Word.Range range7 = paragraph13.Range;
                        range7.Text = "Contract Name\tCredit Shield Parameter\tCredit Shield Value\n";
                        //Orignal query
                        //List_of_dictionaries.c.CommandText = "select 1,'Charge Premium For All Contracts' Type ,t1.name,value,decode(value,1,'YES','NO') Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDUSE' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select  2,'Charge Premium Only For Contracts With Used Credit' ,t1.name,value,decode(value,1,'YES','NO') Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch(+) and T1.TYPE=T2.CONTRACTTYPE(+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDUSEDCL' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select  3,'Domestic Calculation Method',t1.name,value,decode(value,1,'Static',2,'Dynamic',3,'Static With a Post Condition','Static') Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDCALCTYPEDOM' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select 4,'Domestic Calculation Type',t1.name,value,decode(value,1,'Not Used',2,'Charge As Total Of',3,'Charge As Maximum Between',4,'Charge As Minimum Between','Not Used') Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDTYPEDOM' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select 5,'Domestic Amount',t1.name,value,decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDAMOUNTDOM' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select  6,'Domestic Percentage',t1.name,value,decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDPRCDOM' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select  7,'International Calculation Method',t1.name,value,decode(value,1,'Static',2,'Dynamic',3,'Static With a Post Condition','Static') Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDCALCTYPEINT' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select 8,'International Calculation Type',t1.name,value,decode(value,1,'Not Used',2,'Charge As Total Of',3,'Charge As Maximum Between',4,'Charge As Minimum Between','Not Used') Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDTYPEINT' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select  9,'International Amount',t1.name,value,decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDAMOUNTINT' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select  10,'International Percentage',t1.name,value,decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDPRCINT' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select 11,'Promotional Domestic Calculation Type',t1.name,value,decode(value,1,'Not Used',2,'Charge As Total Of',3,'Charge As Maximum Between',4,'Charge As Minimum Between','Not Used') Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDTYPEDOM' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select 12,'Promotional Domestic Amount',t1.name,value,decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDAMOUNTDOM' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select 13,'Promotional Domestic Percentage',t1.name,value,decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDPRCDOM' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select 14,'Promotional Domestic Calendar',t1.name,value,decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDCLNDDOM' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select 15,'Promotional Domestic Cycles Number',t1.name,value,decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDCYCLESDOM' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select 16,'Promotional International Calculation Type',t1.name,value,decode(value,1,'Not Used',2,'Charge As Total Of',3,'Charge As Maximum Between',4,'Charge As Minimum Between','Not Used') Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDTYPEINT' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select 17,'Promotional International Amount',t1.name,value,decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDAMOUNTINT' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select 18,'Promotional International Percentage',t1.name,value,decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDPRCINT' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select 19,'Promotional International Calendar',t1.name,value,decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDCLNDINT' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select 20,'Promotional International Cycles Number',t1.name,value,decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDCYCLESINT' and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " order by 3,1";
                        if (int.Parse(Frm_1.bank_num) == 69)
                        {
                            List_of_dictionaries.c.CommandText = $"select 1,'Charge Premium For All Contracts' Type ,t1.name,value,Decode (VALUE, 2, 'YES',  1, 'YES', 'NO') Output"//decode(value,1,'YES','NO') Output
                        + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                        + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDCHARGEMODE' "//SHIELDUSE 
                        + " and t1.type =" + int.Parse(oracleDataReader03[0].ToString())
                        + " union all select  2,'Charge Premium Only For Contracts With Used Credit' ,t1.name,value,decode(value,2,'YES','NO') Output "
                        + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch(+) and T1.TYPE=T2.CONTRACTTYPE(+) "
                        + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDCHARGEMODE' "//SHIELDUSEDCL
                        + " and t1.type =" + int.Parse(oracleDataReader03[0].ToString())
                        + " union all select  3,'Domestic Calculation Method',t1.name,value,decode(value,1,'Static',2,'Dynamic',3,'Static "
                        + " With a Post Condition','Static') Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 "
                        + " where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch="
                        + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDCALCTYPEDOM' and t1.type ="
                        + int.Parse(oracleDataReader03[0].ToString()) + " union all select 4,'Domestic Calculation Type',t1.name,value, "
                        + " decode(value,1,'Not Used',2,'Charge As Total Of',3,'Charge As Maximum Between',4,'Charge As Minimum Between','Not Used') Output "
                        + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                        + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDTYPEDOM' "
                        + " and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select 5,'Domestic Amount',t1.name,value, "
                        + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                        + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 "
                        + " and t2.key (+)= 'SHIELDAMOUNTDOM' and t1.type =" + int.Parse(oracleDataReader03[0].ToString())
                        + " union all select  6,'Domestic Percentage',t1.name,value,decode(value,null,'Not Used',value) Output "
                        + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                        + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDPRCDOM' and t1.type = "
                        + int.Parse(oracleDataReader03[0].ToString())
                        + " union all select  7,'International Calculation Method',t1.name,value,decode(value,1,'Static',2,'Dynamic',3,'Static "
                        + " With a Post Condition','Static') Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                        + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch= "
                        + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDCALCTYPEINT' and t1.type = "
                        + int.Parse(oracleDataReader03[0].ToString()) + " union all select 8,'International Calculation Type',t1.name,value, "
                        + " decode(value,1,'Not Used',2,'Charge As Total Of',3,'Charge As Maximum Between',4,'Charge As Minimum Between','Not Used') Output "
                        + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                        + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDTYPEINT' and t1.type = "
                        + int.Parse(oracleDataReader03[0].ToString()) + " union all select  9,'International Amount',t1.name,value, "
                        + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                        + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 "
                        + " and t2.key (+)= 'SHIELDAMOUNTINT' and t1.type =" + int.Parse(oracleDataReader03[0].ToString())
                        + " union all select  10,'International Percentage',t1.name,value,decode(value,null,'Not Used',value) Output "
                        + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                        + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDPRCINT' and t1.type ="
                        + int.Parse(oracleDataReader03[0].ToString()) + " union all select 11,'Promotional Domestic Calculation Type',t1.name,value, "
                        + " decode(value,1,'Not Used',2,'Charge As Total Of',3,'Charge As Maximum Between',4,'Charge As Minimum Between','Not Used') Output "
                        + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                        + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDTYPEDOM' and t1.type ="
                        + int.Parse(oracleDataReader03[0].ToString()) + " union all select 12,'Promotional Domestic Amount',t1.name,value, "
                        + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                        + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 "
                        + " and t2.key (+)= 'PROSHIELDAMOUNTDOM' and t1.type =" + int.Parse(oracleDataReader03[0].ToString())
                        + " union all select 13,'Promotional Domestic Percentage',t1.name,value,decode(value,null,'Not Used',value) Output "
                        + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                        + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDPRCDOM' and t1.type ="
                        + int.Parse(oracleDataReader03[0].ToString()) + " union all select 14,'Promotional Domestic Calendar',t1.name,value, "
                        + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                        + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 "
                        + " and t2.key (+)= 'PROSHIELDCLNDDOM' and t1.type =" + int.Parse(oracleDataReader03[0].ToString())
                        + " union all select 15,'Promotional Domestic Cycles Number',t1.name,value,decode(value,null,'Not Used',value) Output "
                        + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                        + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDCYCLESDOM' and t1.type ="
                        + int.Parse(oracleDataReader03[0].ToString()) + " union all select 16,'Promotional International Calculation Type',t1.name,value, "
                        + " decode(value,1,'Not Used',2,'Charge As Total Of',3,'Charge As Maximum Between',4,'Charge As Minimum Between','Not Used') Output "
                        + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                        + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDTYPEINT' and t1.type ="
                        + int.Parse(oracleDataReader03[0].ToString()) + " union all select 17,'Promotional International Amount',t1.name,value, "
                        + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 "
                        + " where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch="
                        + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDAMOUNTINT' and t1.type ="
                        + int.Parse(oracleDataReader03[0].ToString()) + " union all select 18,'Promotional International Percentage',t1.name,value, "
                        + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                        + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch="
                        + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDPRCINT' and t1.type ="
                        + int.Parse(oracleDataReader03[0].ToString()) + " union all select 19,'Promotional International Calendar',t1.name,value, "
                        + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                        + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch="
                        + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDCLNDINT' and t1.type ="
                        + int.Parse(oracleDataReader03[0].ToString()) +
                        " union all " +
                        " select 20,'Promotional International Cycles Number',t1.name,value, "
                        + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                        + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch="
                        + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDCYCLESINT' and t1.type ="
                        + int.Parse(oracleDataReader03[0].ToString()) +
                        " union all " +
                        "SELECT 21, 'When To Charge' Type, t1.name, VALUE, DECODE (VALUE, 1,'Not Used', 2,'Last business day of the month', 3,'Every Statement Date', 4,'Last business day of quarter', 5,'Last Statement Date of quarter', 6,'Every month on specific day', 7,'Specific day of the QRs last month', 8,'First month after the last month of the QR','Not Used') Output " +
                        "FROM a4m.tcontracttype t1, a4m.tcontracttypeparameters t2 " +
                        "WHERE t1.branch = t2.branch(+) AND T1.TYPE = T2.CONTRACTTYPE(+) " +
                        $"AND T1.STATUS = '1' AND t1.branch = {Frm_1.bank_num} AND T1.SCHEMATYPE = 1 " +
                        $"AND t2.key(+) = 'WHENTOCHARGE' AND t1.TYPE = {int.Parse(oracleDataReader03[0].ToString())}" +
                        " union all " +
                        "SELECT 22, 'Base Amount' Type, t1.name, VALUE, DECODE (VALUE, 1,'Current Balance', 2,'Previous Month Balance', 3,'Last Statement Date Balance', 4,'Previous Business Day Balance','Not Used') Output " +
                        "FROM a4m.tcontracttype t1, a4m.tcontracttypeparameters t2 WHERE t1.branch = t2.branch(+) " +
                        $"AND T1.TYPE = T2.CONTRACTTYPE(+) AND T1.STATUS = '1' AND t1.branch = {Frm_1.bank_num} " +
                        $"AND T1.SCHEMATYPE = 1 AND t2.key(+) = 'SHIELDAMOUNTDOM' AND t1.TYPE = {int.Parse(oracleDataReader03[0].ToString())}" +
                        " union all " +
                        "SELECT 23, 'Calculation Method' Type, t1.name, VALUE, DECODE (VALUE, 1,'Based On Fixed Values', 2,'Based On Credit Limit Range','Not Used') Output " +
                        "FROM a4m.tcontracttype t1, a4m.tcontracttypeparameters t2 " +
                        "WHERE t1.branch = t2.branch(+) AND T1.TYPE = T2.CONTRACTTYPE(+) " +
                        $"AND T1.STATUS = '1' AND t1.branch = {Frm_1.bank_num} AND T1.SCHEMATYPE = 1 " +
                        $"AND t2.key(+) = 'SHIELDCALCMETHODDOM' AND t1.TYPE = {int.Parse(oracleDataReader03[0].ToString())}" +
                        " order by 3,1";
                        }
                        else
                        {
                            List_of_dictionaries.c.CommandText = $"select 1,'Charge Premium For All Contracts' Type ,t1.name,value,Decode (VALUE, 2, 'YES',  1, 'YES', 'NO') Output"//decode(value,1,'YES','NO') Output
                                + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                                + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDCHARGEMODE' "//SHIELDUSE 
                                + " and t1.type =" + int.Parse(oracleDataReader03[0].ToString())
                                + " union all select  2,'Charge Premium Only For Contracts With Used Credit' ,t1.name,value,decode(value,2,'YES','NO') Output "
                                + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch(+) and T1.TYPE=T2.CONTRACTTYPE(+) "
                                + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDCHARGEMODE' "//SHIELDUSEDCL
                                + " and t1.type =" + int.Parse(oracleDataReader03[0].ToString())
                                + " union all select  3,'Domestic Calculation Method',t1.name,value,decode(value,1,'Static',2,'Dynamic',3,'Static "
                                + " With a Post Condition','Static') Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 "
                                + " where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch="
                                + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDCALCTYPEDOM' and t1.type ="
                                + int.Parse(oracleDataReader03[0].ToString()) + " union all select 4,'Domestic Calculation Type',t1.name,value, "
                                + " decode(value,1,'Not Used',2,'Charge As Total Of',3,'Charge As Maximum Between',4,'Charge As Minimum Between','Not Used') Output "
                                + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                                + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDTYPEDOM' "
                                + " and t1.type =" + int.Parse(oracleDataReader03[0].ToString()) + " union all select 5,'Domestic Amount',t1.name,value, "
                                + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                                + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 "
                                + " and t2.key (+)= 'SHIELDAMOUNTDOM' and t1.type =" + int.Parse(oracleDataReader03[0].ToString())
                                + " union all select  6,'Domestic Percentage',t1.name,value,decode(value,null,'Not Used',value) Output "
                                + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                                + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDPRCDOM' and t1.type = "
                                + int.Parse(oracleDataReader03[0].ToString())
                                + " union all select  7,'International Calculation Method',t1.name,value,decode(value,1,'Static',2,'Dynamic',3,'Static "
                                + " With a Post Condition','Static') Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                                + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch= "
                                + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDCALCTYPEINT' and t1.type = "
                                + int.Parse(oracleDataReader03[0].ToString()) + " union all select 8,'International Calculation Type',t1.name,value, "
                                + " decode(value,1,'Not Used',2,'Charge As Total Of',3,'Charge As Maximum Between',4,'Charge As Minimum Between','Not Used') Output "
                                + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                                + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDTYPEINT' and t1.type = "
                                + int.Parse(oracleDataReader03[0].ToString()) + " union all select  9,'International Amount',t1.name,value, "
                                + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                                + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 "
                                + " and t2.key (+)= 'SHIELDAMOUNTINT' and t1.type =" + int.Parse(oracleDataReader03[0].ToString())
                                + " union all select  10,'International Percentage',t1.name,value,decode(value,null,'Not Used',value) Output "
                                + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                                + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'SHIELDPRCINT' and t1.type ="
                                + int.Parse(oracleDataReader03[0].ToString()) + " union all select 11,'Promotional Domestic Calculation Type',t1.name,value, "
                                + " decode(value,1,'Not Used',2,'Charge As Total Of',3,'Charge As Maximum Between',4,'Charge As Minimum Between','Not Used') Output "
                                + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                                + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDTYPEDOM' and t1.type ="
                                + int.Parse(oracleDataReader03[0].ToString()) + " union all select 12,'Promotional Domestic Amount',t1.name,value, "
                                + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                                + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 "
                                + " and t2.key (+)= 'PROSHIELDAMOUNTDOM' and t1.type =" + int.Parse(oracleDataReader03[0].ToString())
                                + " union all select 13,'Promotional Domestic Percentage',t1.name,value,decode(value,null,'Not Used',value) Output "
                                + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                                + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDPRCDOM' and t1.type ="
                                + int.Parse(oracleDataReader03[0].ToString()) + " union all select 14,'Promotional Domestic Calendar',t1.name,value, "
                                + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                                + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 "
                                + " and t2.key (+)= 'PROSHIELDCLNDDOM' and t1.type =" + int.Parse(oracleDataReader03[0].ToString())
                                + " union all select 15,'Promotional Domestic Cycles Number',t1.name,value,decode(value,null,'Not Used',value) Output "
                                + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                                + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDCYCLESDOM' and t1.type ="
                                + int.Parse(oracleDataReader03[0].ToString()) + " union all select 16,'Promotional International Calculation Type',t1.name,value, "
                                + " decode(value,1,'Not Used',2,'Charge As Total Of',3,'Charge As Maximum Between',4,'Charge As Minimum Between','Not Used') Output "
                                + " from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) "
                                + " and T1.STATUS='1' and t1.branch=" + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDTYPEINT' and t1.type ="
                                + int.Parse(oracleDataReader03[0].ToString()) + " union all select 17,'Promotional International Amount',t1.name,value, "
                                + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 "
                                + " where t1.branch=t2.branch (+) and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch="
                                + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDAMOUNTINT' and t1.type ="
                                + int.Parse(oracleDataReader03[0].ToString()) + " union all select 18,'Promotional International Percentage',t1.name,value, "
                                + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                                + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch="
                                + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDPRCINT' and t1.type ="
                                + int.Parse(oracleDataReader03[0].ToString()) + " union all select 19,'Promotional International Calendar',t1.name,value, "
                                + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                                + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch="
                                + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDCLNDINT' and t1.type ="
                                + int.Parse(oracleDataReader03[0].ToString()) +
                                " union all " +
                                " select 20,'Promotional International Cycles Number',t1.name,value, "
                                + " decode(value,null,'Not Used',value) Output from a4m.tcontracttype t1,a4m.tcontracttypeparameters t2 where t1.branch=t2.branch (+) "
                                + " and T1.TYPE=T2.CONTRACTTYPE (+) and T1.STATUS='1' and t1.branch="
                                + Frm_1.bank_num + " and T1.SCHEMATYPE=1 and t2.key (+)= 'PROSHIELDCYCLESINT' and t1.type ="
                                + int.Parse(oracleDataReader03[0].ToString()) +
                                " union all " +
                                "SELECT 21, 'When To Charge' Type, t1.name, VALUE, DECODE (VALUE, 1,'Not Used', 2,'Last business day of the month', 3,'Every Statement Date', 4,'Last business day of quarter', 5,'Last Statement Date of quarter', 6,'Every month on specific day', 7,'Specific day of the QRs last month', 8,'First month after the last month of the QR','Not Used') Output " +
                                "FROM a4m.tcontracttype t1, a4m.tcontracttypeparameters t2 " +
                                "WHERE t1.branch = t2.branch(+) AND T1.TYPE = T2.CONTRACTTYPE(+) " +
                                $"AND T1.STATUS = '1' AND t1.branch = {Frm_1.bank_num} AND T1.SCHEMATYPE = 1 " +
                                $"AND t2.key(+) = 'WHENTOCHARGE' AND t1.TYPE = {int.Parse(oracleDataReader03[0].ToString())}" +
                                " union all " +
                                "SELECT 22, 'Base Amount' Type, t1.name, VALUE, DECODE (VALUE, 1,'Current Balance', 2,'Previous Month Balance', 3,'Last Statement Date Balance', 4,'Previous Business Day Balance','Not Used') Output " +
                                "FROM a4m.tcontracttype t1, a4m.tcontracttypeparameters t2 WHERE t1.branch = t2.branch(+) " +
                                $"AND T1.TYPE = T2.CONTRACTTYPE(+) AND T1.STATUS = '1' AND t1.branch = {Frm_1.bank_num} " +
                                $"AND T1.SCHEMATYPE = 1 AND t2.key(+) = 'SHIELDAMOUNTDOM' AND t1.TYPE = {int.Parse(oracleDataReader03[0].ToString())}" +
                                " union all " +
                                "SELECT 23, 'Calculation Method' Type, t1.name, VALUE, DECODE (VALUE, 1,'Based On Fixed Values', 2,'Based On Credit Limit Range','Not Used') Output " +
                                "FROM a4m.tcontracttype t1, a4m.tcontracttypeparameters t2 " +
                                "WHERE t1.branch = t2.branch(+) AND T1.TYPE = T2.CONTRACTTYPE(+) " +
                                $"AND T1.STATUS = '1' AND t1.branch = {Frm_1.bank_num} AND T1.SCHEMATYPE = 1 " +
                                $"AND t2.key(+) = 'SHIELDCALCMETHODDOM' AND t1.TYPE = {int.Parse(oracleDataReader03[0].ToString())}" +
                                " order by 3,1";
                        }
                        //Altered Query 2020-17-02 you will find it in old dectionary version by date 2019-12-04.


                        OracleDataReader oracleDataReader6 = List_of_dictionaries.c.ExecuteReader();
                        while (oracleDataReader6.Read())
                        {
                            Microsoft.Office.Interop.Word.Range range02 = range7;
                            string str = range02.Text + oracleDataReader6[2] + "\t" + oracleDataReader6[1] + "\t" + oracleDataReader6[4] + "\n";
                            range02.Text = str;
                        }
                        oracleDataReader6.Close();
                        range7.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                        range7.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                        range7.Tables[1].Borders.Enable = 1;
                        oPara1 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    }

                }
                if (flag11)
                {
                    oPara1.Range.Text = "\fSection 11 A: Minimum Payment - Credit Products Only\n";
                    oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara1.Range.Font.Bold = 0;
                    oPara1.Range.Font.Size = 9f;
                    oPara1.Format.SpaceAfter = 0.0f;
                    oPara1.Range.InsertParagraphAfter();
                    Paragraph paragraph2 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph2.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range1 = paragraph2.Range;
                    range1.Text = "Profile Name\tProfile Value\tProfile Percentage\tProfile Type\n";
                    List_of_dictionaries.c.CommandText = "select v0 Profile,sum(v1) Amount,sum(v2) Percentage,decode(sum(v3),1,'As Total Of',2,'As Maximum Between',3,'As Minimum Between') Type from (select distinct tcontractmpprofile.PROFILENAME v0,tcontractotherparameters.VALUE v1,null v2 ,null v3 from a4m.tcontractotherparameters,a4m.tcontractmpprofile,a4m.tcontractotherobjecttype where tcontractotherparameters.branch=" + Frm_1.bank_num + " and Upper(tcontractotherparameters.KEY) like 'MINPAYAMOUNT' and tcontractmpprofile.branch=tcontractotherparameters.branch and tcontractmpprofile.PROFILEID=tcontractotherparameters.OBJECTNO and tcontractotherobjecttype.branch=tcontractotherparameters.branch and tcontractotherobjecttype.OBJECTTYPE=tcontractotherparameters.OBJECTTYPE union all select distinct tcontractmpprofile.PROFILENAME v0,null v1,tcontractotherparameters.VALUE v2,null v3 from a4m.tcontractotherparameters,a4m.tcontractmpprofile,a4m.tcontractotherobjecttype where tcontractotherparameters.branch=" + Frm_1.bank_num + " and Upper(tcontractotherparameters.KEY) like 'MINPAYPRC'and tcontractmpprofile.branch=tcontractotherparameters.branch and tcontractmpprofile.PROFILEID=tcontractotherparameters.OBJECTNO and tcontractotherobjecttype.branch=tcontractotherparameters.branch and tcontractotherobjecttype.OBJECTTYPE=tcontractotherparameters.OBJECTTYPE union all select distinct tcontractmpprofile.PROFILENAME v0,null v1,null v2, tcontractotherparameters.VALUE v3 from a4m.tcontractotherparameters,a4m.tcontractmpprofile,a4m.tcontractotherobjecttype where tcontractotherparameters.branch=" + Frm_1.bank_num + " and Upper(tcontractotherparameters.KEY) like 'MINPAYTYPE' and tcontractmpprofile.branch=tcontractotherparameters.branch and tcontractmpprofile.PROFILEID=tcontractotherparameters.OBJECTNO and tcontractotherobjecttype.branch=tcontractotherparameters.branch and tcontractotherobjecttype.OBJECTTYPE=tcontractotherparameters.OBJECTTYPE )group by v0 order by 1";
                    OracleDataReader oracleDataReader1 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader1.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range1;
                        string str = range2.Text + oracleDataReader1[0] + "\t" + oracleDataReader1[1] + "\t" + oracleDataReader1[2] + "\t" + oracleDataReader1[3] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader1.Close();
                    range1.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range1.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range1.Tables[1].Borders.Enable = 1;
                    paragraph12 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    Paragraph paragraph3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph3.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range3 = paragraph3.Range;
                    range3.Text = "Contract Name\tContract Profile\tContract Profile Currency\n";
                    List_of_dictionaries.c.CommandText = "select tcontracttype.name,tcontractmpprofile.PROFILENAME ,'Domestic' Currency from a4m.tcontractmpprofile,a4m.tcontracttype,a4m.tcontracttypeparameters where tcontracttype.branch=tcontracttypeparameters.branch and tcontracttype.TYPE=tcontracttypeparameters.CONTRACTTYPE and tcontracttypeparameters.KEY like 'MPPROFILEDOM'and tcontracttypeparameters.branch=tcontractmpprofile.branch and tcontracttypeparameters.VALUE=tcontractmpprofile.PROFILEID and tcontracttype.branch=" + Frm_1.bank_num + " and tcontracttype.status like '1' union all select tcontracttype.name,tcontractmpprofile.PROFILENAME ,'International' Currency from a4m.tcontractmpprofile,a4m.tcontracttype,a4m.tcontracttypeparameters where tcontracttype.branch=tcontracttypeparameters.branch and tcontracttype.TYPE=tcontracttypeparameters.CONTRACTTYPE and tcontracttypeparameters.KEY like 'MPPROFILEINT' and tcontracttypeparameters.branch=tcontractmpprofile.branch and tcontracttypeparameters.VALUE=tcontractmpprofile.PROFILEID and tcontracttype.branch=" + Frm_1.bank_num + " and tcontracttype.status like '1' order by 1";
                    OracleDataReader oracleDataReader2 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader2.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range3;
                        string str = range2.Text + oracleDataReader2[0] + "\t" + oracleDataReader2[1] + "\t" + oracleDataReader2[2] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader2.Close();
                    range3.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range3.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range3.Tables[1].Borders.Enable = 1;
                    Paragraph paragraph4 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph4.Range.Text = "\fSection 11 B : Direct Debit - Credit Products Only\n";
                    paragraph4.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    paragraph4.Range.Font.Bold = 0;
                    paragraph4.Range.Font.Size = 9f;
                    paragraph4.Format.SpaceAfter = 6f;
                    paragraph4.Range.InsertParagraphAfter();
                    Paragraph paragraph5 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph5.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range4 = paragraph5.Range;
                    range4.Text = "DD Profile Name\tDD Profile Type\tDD Calculation Method\tDD Profile Value\tDD Profile Percentage\tDD Profile Based On\n";
                    List_of_dictionaries.c.CommandText = "select d.PROFILENAME,decode(PLSQLID,0,'STATIC','DYNAMIC') Type,decode(d.DDMETHOD,1,'As Total Of',2,'As Maximum Between',3,'As minimum Between')  method,d.DDAMOUNT,d.DDPRC,decode(d.DDBASEDON,1,'Total Outstanding Amount',2,'Minimum Payment Amount',3,'Overdue Amount')  Base from a4m.tcontractddreference d where branch=" + Frm_1.bank_num + " order by 1";
                    OracleDataReader oracleDataReader3 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader3.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range4;
                        string str = range2.Text + oracleDataReader3[0] + "\t" + oracleDataReader3[1] + "\t" + oracleDataReader3[2] + "\t" + oracleDataReader3[3] + "\t" + oracleDataReader3[4] + "\t" + oracleDataReader3[5] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader3.Close();
                    range4.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range4.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range4.Tables[1].Borders.Enable = 1;
                    paragraph12 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    Paragraph paragraph6 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph6.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range5 = paragraph6.Range;
                    range5.Text = "Contract Name\tContract DD Profile\tContract DD Profile Currency\n";
                    List_of_dictionaries.c.CommandText = "select tcontracttype.name,tcontractddreference.PROFILENAME ,'Domestic' Currency from a4m.tcontractddreference,a4m.tcontracttype,a4m.tcontracttypeddsettings where tcontracttype.branch=tcontracttypeddsettings.branch and tcontracttype.TYPE=tcontracttypeddsettings.CONTRACTTYPE and tcontracttypeddsettings.KEY like 'DAFPROFILEDOM' and tcontracttypeddsettings.branch=tcontractddreference.branch and tcontracttypeddsettings.PROFILEID=tcontractddreference.PROFILEID and tcontracttype.branch=" + Frm_1.bank_num + " and tcontracttype.status like '1' union all select tcontracttype.name,tcontractddreference.PROFILENAME ,'International' Currency from a4m.tcontractddreference,a4m.tcontracttype,a4m.tcontracttypeddsettings where tcontracttype.branch=tcontracttypeddsettings.branch and tcontracttype.TYPE=tcontracttypeddsettings.CONTRACTTYPE and tcontracttypeddsettings.KEY like 'DAFPROFILEINT' and tcontracttypeddsettings.branch=tcontractddreference.branch and tcontracttypeddsettings.PROFILEID=tcontractddreference.PROFILEID and tcontracttype.branch=" + Frm_1.bank_num + " and tcontracttype.status like '1' order by 1";
                    OracleDataReader oracleDataReader4 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader4.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range5;
                        string str = range2.Text + oracleDataReader4[0] + "\t" + oracleDataReader4[1] + "\t" + oracleDataReader4[2] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader4.Close();
                    range5.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range5.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range5.Tables[1].Borders.Enable = 1;
                    oPara1 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                }
                if (flag12)
                {
                    oPara1.Range.Text = "\fSection 12 : Card Limits - Credit Products Only\n";
                    oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara1.Range.Font.Bold = 0;
                    oPara1.Range.Font.Size = 9f;
                    oPara1.Format.SpaceAfter = 0.0f;
                    oPara1.Range.InsertParagraphAfter();
                    Paragraph paragraph2 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);

                    //edt 1002  samr   sort primary and Supplementary sections

                    paragraph2.Range.Text = "S12:1-Account Default Withdrawal Limit\n";
                    paragraph2.Range.InsertParagraphAfter();
                    Paragraph paragraph3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph3.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range1 = paragraph3.Range;
                    range1.Text = "Contract Type\tFlat Amount\tPercentage\tCalculation Type\tCurrency\n";
                    //List_of_dictionaries.c.CommandText = "select name,sum(amount),sum(prc),decode(sum(type),1,'As Total Of',2,'As Maximum Of',3,'As Minimum Of','Not Used (100%)')  Type,c currency from (select t2.NAME name,t1.VALUE amount,null prc,'International' c,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'ACCCASHLIMITAMOUNTINT'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,null prc,'International' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'ACCCASHLIMITTYPEINT' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'International' c,null Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "  and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'ACCCASHLIMITPRCINT' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,t1.VALUE amount,null prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'ACCCASHLIMITAMOUNTDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,null prc,'Domestic' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'ACCCASHLIMITTYPEDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "  and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'ACCCASHLIMITPRCDOM' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID)group by name ,c order by 1";

                    //UBA-1525 9-6-2020 msattar adding and t4.currencyno to every where -> currnecyno = 2 for International queries and currnecyno = 1 for Domastic queries
                    //and from every query in select  union have 't4.perecent Type' will be  't4.calcmethod Type' to fix the decode issues. as 1 = total of , 2=Maximum between , 3 = minimum between , 0 =Not Used(100%)
                    //Note as example:if t4.limitkind = 1,t4.objecttype =2, t4.ownership = 0,t4.limitid =1002 => means get Account defualt withdrowl limit and so on it varies depend on these 4 parameters
                    //every section in 12.1 to 12.7 have 6 unions 3 queries for international and 3 queries for DOMISTIC , each of the 3 
                    //have identifier first one gets amount second gets type"calculated method", third gets percentage
                    List_of_dictionaries.c.CommandText =
                        " select name,sum(amount),sum(prc),decode(sum(type),1,'As Total Of',2,'As Maximum Of',3,'As Minimum Of','Not Used (100%)')  Type, "
+ " c currency "
+ " from (  "
+ "                     select t2.NAME name,t4.Flatamount amount,null prc,'International' c, null type  "
+ "                   from  a4m.tcontracttypeparameters t1, "
+ "                           a4m.tcontracttype t2, "
+ "                           a4m.tcontractprofile t3, "
+ "                           a4m.TCONTRACTTYPELIMITSSETTINGS t4 "
+ "                    where t1.branch= '" + Frm_1.bank_num + "' "
+ "                     and t4.branch = t1.branch "
+ "                     and t4.contracttype = t1.contracttype "
+ "                     and t1.BRANCH=t2.BRANCH "
+ "                     and t1.CONTRACTTYPE=t2.TYPE "
+ "                    and (upper(t1.KEY)like'ACCCASHLIMITAMOUNTINT' OR upper(t1.KEY) Like 'PROFILEINT') "
+ "                     and t1.branch=t3.branch "
+ "                     and t1.VALUE=t3.PROFILEID  "
+ "                     and t2.status like '1' "
+ "                    and t4.currencyno=2 and t4.limitkind = 1 and t4.objecttype = 2 and t4.ownership = 0 and t4.limitid = 1002 "

+ "                     union all "

+ "                     select t2.NAME name,null amount,null prc,'International' c ,t4.calcmethod Type "
+ "                     from a4m.tcontracttypeparameters t1, "
+ "                          a4m.tcontracttype t2, "
+ "                          a4m.tcontractprofile t3, "
+ "                          a4m.TCONTRACTTYPELIMITSSETTINGS t4 "
+ "                    where t1.branch='" + Frm_1.bank_num + "' "
+ "                    and t4.branch = t1.branch "
+ "                    and t4.contracttype = t1.contracttype "
+ "                    and t1.BRANCH=t2.BRANCH "
+ "                    and t1.CONTRACTTYPE=t2.TYPE "
+ "                    and (upper(t1.KEY)like'ACCCASHLIMITTYPEINT' OR upper(t1.KEY) Like 'PROFILEINT') "
+ "                    and t1.branch=t3.branch "
+ "                    and t1.value=t3.PROFILEID "
+ "                    and t2.status like '1' "
+ "                    and t4.currencyno=2  and t4.limitkind = 1 and t4.objecttype = 2 and t4.ownership = 0 and t4.limitid =1002 "

+ "                    union all "

+ "                    select t2.NAME name,null amount,t4.percent prc,'International' c,null Type "
+ "                    from a4m.tcontracttypeparameters t1, "
+ "                        a4m.tcontracttype t2, "
+ "                        a4m.tcontractprofile t3, "
+ "                       a4m.TCONTRACTTYPELIMITSSETTINGS t4 "

+ "                    where "
+ "                    t1.branch='" + Frm_1.bank_num + "' "
+ "                    and t4.branch = t1.branch "
+ "                    and t4.contracttype = t1.contracttype "
+ "                    and t1.BRANCH=t2.BRANCH "
+ "                    and t1.CONTRACTTYPE=t2.TYPE "
+ "                    and (upper(t1.KEY)like'ACCCASHLIMITPRCINT' OR upper(t1.KEY) Like 'PROFILEINT') "
+ "                    and t1.branch=t3.branch "
+ "                    and t1.value=t3.PROFILEID "
+ "                    and t2.status like '1' "
+ "                    and t4.currencyno=2  and t4.limitkind = 1 and t4.objecttype = 2 and t4.ownership = 0 and t4.limitid =1002 "

+ "                    union all "

+ "                    select t2.NAME name,t4.Flatamount amount,null prc,'Domestic' c ,null type "
+ "                    from a4m.tcontracttypeparameters t1, "
+ "                        a4m.tcontracttype t2,"
+ "                        a4m.tcontractprofile t3, "
+ "                        a4m.TCONTRACTTYPELIMITSSETTINGS t4 "
+ "                    where t1.branch='" + Frm_1.bank_num + "' "
+ "                    and t4.branch = t1.branch "
+ "                    and t4.contracttype = t1.contracttype "
+ "                    and t1.BRANCH=t2.BRANCH "
+ "                    and t1.CONTRACTTYPE=t2.TYPE "
+ "                    and (upper(t1.KEY)like'ACCCASHLIMITAMOUNTDOM' OR upper(t1.KEY) Like 'PROFILEDOM' ) "
+ "                    and t1.branch=t3.branch "
+ "                    and t1.value=t3.PROFILEID "
+ "                    and t2.status like '1'   "
+ "                    and t4.currencyno=1  and t4.limitkind = 1 and t4.objecttype = 2 and t4.ownership = 0 and t4.limitid =1002 "

+ "                    union all "

+ "                    select t2.NAME name,null amount,null prc,'Domestic' c ,t4.calcmethod Type "
+ "                    from a4m.tcontracttypeparameters t1, "
+ "                         a4m.tcontracttype t2, "
+ "                        a4m.tcontractprofile t3 , "
+ "                        a4m.TCONTRACTTYPELIMITSSETTINGS t4 "

+ "                    where t1.branch='" + Frm_1.bank_num + "' "
+ "                    and t4.branch = t1.branch "
+ "                    and t4.contracttype = t1.contracttype "
+ "                    and t1.BRANCH=t2.BRANCH  "
+ "                    and t1.CONTRACTTYPE=t2.TYPE "
+ "                    and (upper(t1.KEY)like'ACCCASHLIMITTYPEDOM' or upper(t1.KEY) Like 'PROFILEDOM') "
+ "                    and t1.branch=t3.branch "
+ "                    and t1.value=t3.PROFILEID "
+ "                    and t2.status like '1' "
+ "                    and t4.currencyno=1  and t4.limitkind = 1 and t4.objecttype = 2 and t4.ownership = 0 and t4.limitid =1002 "

+ "                    union all "

+ "                    select t2.NAME name,null amount,t4.percent prc,'Domestic' c ,null type "
+ "                    from a4m.tcontracttypeparameters t1, "
+ "                         a4m.tcontracttype t2, "
+ "                         a4m.tcontractprofile t3 , "
+ "                         a4m.TCONTRACTTYPELIMITSSETTINGS t4 "

+ "                    where t1.branch= '" + Frm_1.bank_num + "' "
+ "                    and t4.branch = t1.branch "
+ "                    and t4.contracttype = t1.contracttype "
+ "                    and t1.BRANCH=t2.BRANCH "
+ "                    and t1.CONTRACTTYPE=t2.TYPE "
+ "                    and (upper(t1.KEY)like'ACCCASHLIMITPRCDOM' or upper(t1.KEY) Like 'PROFILEDOM') "
+ "                    and t1.branch=t3.branch "
+ "                    and t1.value=t3.PROFILEID "
+ "                    and t4.currencyno=1  and t4.limitkind = 1 and t4.objecttype = 2 and t4.ownership = 0 and t4.limitid =1002 ) "

+ " /*put this here if you want one contract  where lower(NAme)like lower('VISA GOLD SECURED%') */ "

 + "                    group by name ,c order by 1 ";
                    OracleDataReader oracleDataReader1 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader1.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range1;
                        string str = range2.Text + oracleDataReader1[0] + "\t" + oracleDataReader1[1] + "\t" + oracleDataReader1[2] + "\t" + oracleDataReader1[3] + "\t" + oracleDataReader1[4] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader1.Close();
                    range1.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range1.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range1.Tables[1].Borders.Enable = 1;
                    Paragraph paragraph4 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph4.Range.Text = "\nS12:2-Primary Card Default Credit Limit\n";
                    paragraph4.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph4.Range.InsertParagraphAfter();
                    Paragraph paragraph5 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph5.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range3 = paragraph5.Range;
                    range3.Text = "Contract Type\tFlat Amount\tPercentage\tCalculation Type\tCurrency\n";
                    //List_of_dictionaries.c.CommandText = "select name,sum(amount),sum(prc),decode(sum(type),1,'As Total Of',2,'As Maximum Of',3,'As Minimum Of','Not Used (100%)')  Type,c currency from (select t2.NAME name,t1.VALUE amount,null prc,'International' c,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile  where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICREDLIMITAMOUNTINT'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch  and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,null prc,'International' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICREDLIMITTYPEINT'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'International' c,null Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICREDLIMITPRCINT' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,t1.VALUE amount,null prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICREDLIMITAMOUNTDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID  and t2.status like '1' union all select t2.NAME name,null amount,null prc,'Domestic' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile  where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICREDLIMITTYPEDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICREDLIMITPRCDOM' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID) group by name ,c order by 1";

                    //altered in tables and where statment as 12.1 but with changes in where condition in last 4 parameters
                    List_of_dictionaries.c.CommandText = @"/* Formatted on 3/15/2021 11:52:09 AM (QP5 v5.227.12220.39754) */
--12.2
  SELECT name,
         SUM (amount),
         SUM (prc),
         DECODE (SUM (TYPE),
                 1, 'As Total Of',
                 2, 'As Maximum Of',
                 3, 'As Minimum Of',
                 'Not Used (100%)')
            TYPE,
         c currency
    FROM (SELECT t2.NAME name,
                 t4.Flatamount amount,
                 NULL prc,
                 'International' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND (   UPPER (t1.KEY) LIKE 'PRICREDLIMITAMOUNTINT' 
                      OR UPPER (t1.KEY) LIKE 'PROFILEINT')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 2
                 and t4.limitkind = 0
                 and t4.objecttype =3
                 and t4.ownership = 0
                 and t4.limitid =1001 
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 NULL prc,
                 'International' c,
                 t4.calcmethod TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND (   UPPER (t1.KEY) LIKE 'PRICREDLIMITTYPEINT' 
                      OR UPPER (t1.KEY) LIKE 'PROFILEINT')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 2
                 and t4.limitkind = 0
                 and t4.objecttype =3
                 and t4.ownership = 0
                 and t4.limitid =1001 
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 t4.percent prc,
                 'International' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND (   UPPER (t1.KEY) LIKE 'PRICREDLIMITPRCINT' 
                      OR UPPER (t1.KEY) LIKE 'PROFILEINT')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 2
                 and t4.limitkind = 0
                 and t4.objecttype =3
                 and t4.ownership = 0
                 and t4.limitid =1001 
          UNION ALL
          SELECT t2.NAME name,
                 t4.Flatamount amount,
                 NULL prc,
                 'Domestic' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND ( UPPER (t1.KEY) LIKE 'PRICREDLIMITAMOUNTDOM'
                     OR UPPER (t1.KEY) LIKE 'PROFILEDOM')
                  AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 1 
                 and t4.limitkind = 0
                 and t4.objecttype =3
                 and t4.ownership = 0
                 and t4.limitid =1001
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 NULL prc,
                 'Domestic' c,
                 t4.calcmethod TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND ( UPPER (t1.KEY) LIKE 'PRICREDLIMITTYPEDOM'
                 OR UPPER (t1.KEY) LIKE 'PROFILEDOM')
                  AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 1 
                 and t4.limitkind = 0
                 and t4.objecttype =3
                 and t4.ownership = 0
                 and t4.limitid =1001
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 t4.percent prc,
                 'Domestic' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,                 
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND ( UPPER (t1.KEY) LIKE 'PRICREDLIMITPRCDOM'
                 OR UPPER (t1.KEY) LIKE 'PROFILEDOM')
                  AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 1
                 and t4.limitkind = 0
                 and t4.objecttype =3
                 and t4.ownership = 0
                 and t4.limitid =1001 
                 )
GROUP BY name, c
ORDER BY 1";
                    OracleDataReader oracleDataReader2 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader2.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range3;
                        string str = range2.Text + oracleDataReader2[0] + "\t" + oracleDataReader2[1] + "\t" + oracleDataReader2[2] + "\t" + oracleDataReader2[3] + "\t" + oracleDataReader2[4] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader2.Close();
                    range3.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range3.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range3.Tables[1].Borders.Enable = 1;
                    Paragraph paragraph6 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph6.Range.Text = "\nS12:3-Primary Card Default Withdrawal Limit\n";
                    paragraph6.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph6.Range.InsertParagraphAfter();
                    Paragraph paragraph7 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph7.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range4 = paragraph7.Range;
                    range4.Text = "Contract Type\tFlat Amount\tPercentage\tCalculation Type\tCurrency\n";

                    //List_of_dictionaries.c.CommandText = "select name,sum(amount),sum(prc),decode(sum(type),1,'As Total Of',2,'As Maximum Of',3,'As Minimum Of','Not Used (100%)')  Type,c currency from (select t2.NAME name,t1.VALUE amount,null prc,'International' c,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile  where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICASHLIMITAMOUNTINT'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch  and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,null prc,'International' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICASHLIMITTYPEINT'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'International' c,null Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICASHLIMITPRCINT' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,t1.VALUE amount,null prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICASHLIMITAMOUNTDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID  and t2.status like '1' union all select t2.NAME name,null amount,null prc,'Domestic' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile  where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICASHLIMITTYPEDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICASHLIMITPRCDOM' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID) group by name ,c order by 1";

                    //alter same as 12.2 but the differnece in where statments
                    List_of_dictionaries.c.CommandText = @"/* Formatted on 3/15/2021 11:53:09 AM (QP5 v5.227.12220.39754) */
--12.3

  SELECT name,
         SUM (amount),
         SUM (prc),
         DECODE (SUM (TYPE),
                 1, 'As Total Of',
                 2, 'As Maximum Of',
                 3, 'As Minimum Of',
                 'Not Used (100%)')
            TYPE,
         c currency
    FROM (SELECT t2.NAME name,
                 t4.Flatamount amount,
                 NULL prc,
                 'International' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND ( UPPER (t1.KEY) LIKE 'PRICASHLIMITAMOUNTINT'
                      OR UPPER (t1.KEY) LIKE  'PROFILEINT')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 2
                 and t4.limitkind = 1
                 and t4.objecttype =3
                 and t4.ownership = 0
                 and t4.limitid =1002
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 NULL prc,
                 'International' c,
                 t4.calcmethod TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND (UPPER (t1.KEY) LIKE 'PRICASHLIMITTYPEINT'
                 OR UPPER  (t1.KEY) LIKE 'PROFILEINT')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 2
                 and t4.limitkind = 1
                 and t4.objecttype =3
                 and t4.ownership = 0
                 and t4.limitid =1002
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 t4.percent prc,
                 'International' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND (UPPER (t1.KEY) LIKE 'PRICASHLIMITPRCINT'
                 OR  UPPER  (t1.KEY) LIKE 'PROFILEINT')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 2
                 and t4.limitkind = 1
                 and t4.objecttype =3
                 and t4.ownership = 0
                 and t4.limitid =1002
          UNION ALL
          SELECT t2.NAME name,
                 t4.Flatamount amount,
                 NULL prc,
                 'Domestic' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND ( UPPER (t1.KEY) LIKE 'PRICASHLIMITAMOUNTDOM'
                 OR UPPER (t1.KEY) LIKE 'PROFILEDOM')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 1
                 and t4.limitkind = 1
                 and t4.objecttype =3
                 and t4.ownership = 0
                 and t4.limitid =1002
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 NULL prc,
                 'Domestic' c,
                 t4.calcmethod TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND ( UPPER (t1.KEY) LIKE 'PRICASHLIMITTYPEDOM'
                 OR UPPER (t1.KEY) LIKE 'PROFILEDOM')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 1
                 and t4.limitkind = 1
                 and t4.objecttype =3
                 and t4.ownership = 0
                 and t4.limitid =1002
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 t4.percent prc,
                 'Domestic' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,                 
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND ( UPPER (t1.KEY) LIKE 'PRICASHLIMITPRCDOM'
                  OR  UPPER (t1.KEY) LIKE 'PROFILEDOM')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 1
                 and t4.limitkind = 1
                 and t4.objecttype =3
                 and t4.ownership = 0
                 and t4.limitid =1002)
GROUP BY name, c
ORDER BY 1";
                    OracleDataReader oracleDataReader3 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader3.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range4;
                        string str = range2.Text + oracleDataReader3[0] + "\t" + oracleDataReader3[1] + "\t" + oracleDataReader3[2] + "\t" + oracleDataReader3[3] + "\t" + oracleDataReader3[4] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader3.Close();
                    range4.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range4.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range4.Tables[1].Borders.Enable = 1;
                    //EDT-1002
                    Paragraph paragraph8 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph8.Range.Text = "\nS12:4-Supplementary Card Default Credit Limit\n";
                    paragraph8.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph8.Range.InsertParagraphAfter();
                    Paragraph paragraph9 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph9.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range5 = paragraph9.Range;
                    range5.Text = "Contract Type\tFlat Amount\tPercentage\tCalculation Type\tCurrency\n";
                    //List_of_dictionaries.c.CommandText = "select name,sum(amount),sum(prc),decode(sum(type),1,'As Total Of',2,'As Maximum Of',3,'As Minimum Of','Not Used (100%)')  Type,c currency from (select t2.NAME name,t1.VALUE amount,null prc,'International' c,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile  where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCREDLIMITAMOUNTINT'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch  and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,null prc,'International' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCREDLIMITTYPEINT'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'International' c,null Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCREDLIMITPRCINT' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,t1.VALUE amount,null prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCREDLIMITAMOUNTDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID  and t2.status like '1' union all select t2.NAME name,null amount,null prc,'Domestic' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile  where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCREDLIMITTYPEDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCREDLIMITPRCDOM' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID) group by name ,c order by 1";

                    //altered like 12.3 differnce in where
                    List_of_dictionaries.c.CommandText = @"/* Formatted on 3/15/2021 11:53:44 AM (QP5 v5.227.12220.39754) */
--12.4 
 SELECT name,
         SUM (amount),
         SUM (prc),
         DECODE (SUM (TYPE),
                 1, 'As Total Of',
                 2, 'As Maximum Of',
                 3, 'As Minimum Of',
                 'Not Used (100%)')
            TYPE,
         c currency
    FROM (SELECT t2.NAME name,
                 t4.Flatamount amount,
                 NULL prc,
                 'International' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND (UPPER (t1.KEY) LIKE 'SUPCREDLIMITAMOUNTINT'
                 OR UPPER (t1.KEY) LIKE  'PROFILEINT')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 2
                 and t4.limitkind = 0
                 and t4.objecttype =3
                 and t4.ownership = 1
                 and t4.limitid =1001
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 NULL prc,
                 'International' c,
                 t4.calcmethod TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND (UPPER (t1.KEY) LIKE 'SUPCREDLIMITTYPEINT'
                 OR UPPER (t1.KEY) LIKE  'PROFILEINT')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 2
                 and t4.limitkind = 0
                 and t4.objecttype =3
                 and t4.ownership = 1
                 and t4.limitid =1001
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 t4.percent prc,
                 'International' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND (UPPER (t1.KEY) LIKE 'SUPCREDLIMITPRCINT'
                 OR UPPER (t1.KEY) LIKE  'PROFILEINT')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 2
                 and t4.limitkind = 0
                 and t4.objecttype =3
                 and t4.ownership = 1
                 and t4.limitid =1001
          UNION ALL
          SELECT t2.NAME name,
                 t4.Flatamount amount,
                 NULL prc,
                 'Domestic' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND ( UPPER (t1.KEY) LIKE 'SUPCREDLIMITAMOUNTDOM'
                 OR UPPER (t1.KEY) LIKE  'PROFILEDOM')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 1
                 and t4.limitkind = 0
                 and t4.objecttype =3
                 and t4.ownership = 1
                 and t4.limitid =1001
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 NULL prc,
                 'Domestic' c,
                 t4.calcmethod TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND ( UPPER (t1.KEY) LIKE 'SUPCREDLIMITTYPEDOM'
                 OR UPPER (t1.KEY) LIKE  'PROFILEDOM')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 1
                 and t4.limitkind = 0
                 and t4.objecttype =3
                 and t4.ownership = 1
                 and t4.limitid =1001
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 t4.percent prc,
                 'Domestic' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch = " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND ( UPPER (t1.KEY) LIKE 'SUPCREDLIMITPRCDOM'
                 OR UPPER (t1.KEY) LIKE  'PROFILEDOM')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 1
                 and t4.limitkind = 0
                 and t4.objecttype =3
                 and t4.ownership = 1
                 and t4.limitid =1001)
GROUP BY name, c
ORDER BY 1";
                    OracleDataReader oracleDataReader4 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader4.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range5;
                        string str = range2.Text + oracleDataReader4[0] + "\t" + oracleDataReader4[1] + "\t" + oracleDataReader4[2] + "\t" + oracleDataReader4[3] + "\t" + oracleDataReader4[4] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader4.Close();
                    range5.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range5.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range5.Tables[1].Borders.Enable = 1;
                    Paragraph paragraph10 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph10.Range.Text = "\nS12:5-Supplementary Card Default Withdrawal Limit\n";
                    paragraph10.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph10.Range.InsertParagraphAfter();
                    Paragraph paragraph11 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph11.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range6 = paragraph11.Range;
                    range6.Text = "Contract Type\tFlat Amount\tPercentage\tCalculation Type\tCurrency\n";
                    //old query
                    //List_of_dictionaries.c.CommandText = "select name,sum(amount),sum(prc),decode(sum(type),1,'As Total Of',2,'As Maximum Of',3,'As Minimum Of','Not Used (100%)')  Type,c currency from (select t2.NAME name,t1.VALUE amount,null prc,'International' c,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile  where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCASHLIMITAMOUNTINT'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch  and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,null prc,'International' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCASHLIMITTYPEINT'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'International' c,null Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCASHLIMITPRCINT' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,t1.VALUE amount,null prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCASHLIMITAMOUNTDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID  and t2.status like '1' union all select t2.NAME name,null amount,null prc,'Domestic' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile  where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCASHLIMITTYPEDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCASHLIMITPRCDOM' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID) group by name ,c order by 1";

                    //msattar 09082020 like 12.1 we add table TCONTRACTTYPELIMITSSETTINGS and do as 12.1
                    List_of_dictionaries.c.CommandText = @"
  /* Formatted on 3/15/2021 11:53:44 AM (QP5 v5.227.12220.39754) */
--12.5 
 SELECT name,
         SUM (amount),
         SUM (prc),
         DECODE (SUM (TYPE),
                 1, 'As Total Of',
                 2, 'As Maximum Of',
                 3, 'As Minimum Of',
                 'Not Used (100%)')
            TYPE,
         c currency
    FROM (SELECT t2.NAME name,
                 t4.Flatamount amount,
                 NULL prc,
                 'International' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch =  " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND (UPPER (t1.KEY) LIKE 'SUPCASHLIMITAMOUNTINT'
                 OR UPPER (t1.KEY) LIKE  'PROFILEINT')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 2
                 and t4.limitkind = 1
                 and t4.objecttype =3
                 and t4.ownership = 1
                 and t4.limitid =1002
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 NULL prc,
                 'International' c,
                 t4.calcmethod TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch =  " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND (UPPER (t1.KEY) LIKE 'SUPCASHLIMITTYPEINT'
                 OR UPPER (t1.KEY) LIKE  'PROFILEINT')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 2
                 and t4.limitkind = 1
                 and t4.objecttype =3
                 and t4.ownership = 1
                 and t4.limitid =1002
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 t4.percent prc,
                 'International' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch =  " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND (UPPER (t1.KEY) LIKE 'SUPCASHLIMITPRCINT'
                 OR UPPER (t1.KEY) LIKE  'PROFILEINT')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 2
                 and t4.limitkind = 1
                 and t4.objecttype =3
                 and t4.ownership = 1
                 and t4.limitid =1002
          UNION ALL
          SELECT t2.NAME name,
                 t4.Flatamount amount,
                 NULL prc,
                 'Domestic' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch =  " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND ( UPPER (t1.KEY) LIKE 'SUPCASHLIMITAMOUNTDOM'
                 OR UPPER (t1.KEY) LIKE  'PROFILEDOM')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 1
                 and t4.limitkind = 1
                 and t4.objecttype =3
                 and t4.ownership = 1
                 and t4.limitid =1002
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 NULL prc,
                 'Domestic' c,
                 t4.calcmethod TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch =  " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND ( UPPER (t1.KEY) LIKE 'SUPCASHLIMITTYPEDOM'
                 OR UPPER (t1.KEY) LIKE  'PROFILEDOM')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 1
                 and t4.limitkind = 1
                 and t4.objecttype =3
                 and t4.ownership = 1
                 and t4.limitid =1002
          UNION ALL
          SELECT t2.NAME name,
                 NULL amount,
                 t4.percent prc,
                 'Domestic' c,
                 NULL TYPE
            FROM a4m.tcontracttypeparameters t1,
                 a4m.tcontracttype t2,
                 a4m.tcontractprofile t3,
                 a4m.TCONTRACTTYPELIMITSSETTINGS t4
           WHERE     t1.branch =  " + Frm_1.bank_num + @"
                 AND t4.branch = t1.branch
                 AND t4.contracttype = t1.contracttype
                 AND t1.BRANCH = t2.BRANCH
                 AND t1.CONTRACTTYPE = t2.TYPE
                 AND ( UPPER (t1.KEY) LIKE 'SUPCASHLIMITPRCDOM'
                 OR UPPER (t1.KEY) LIKE  'PROFILEDOM')
                 AND t1.branch = t3.branch
                 AND t1.VALUE = t3.PROFILEID
                 AND t2.status LIKE '1'
                 AND t4.currencyno = 1
                 and t4.limitkind = 1
                 and t4.objecttype =3
                 and t4.ownership = 1
                 and t4.limitid =1002)
GROUP BY name, c
ORDER BY 1 ";
                    OracleDataReader oracleDataReader5 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader5.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range6;
                        string str = range2.Text + oracleDataReader5[0] + "\t" + oracleDataReader5[1] + "\t" + oracleDataReader5[2] + "\t" + oracleDataReader5[3] + "\t" + oracleDataReader5[4] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader5.Close();
                    range6.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range6.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range6.Tables[1].Borders.Enable = 1;

                    //iatta
                    //Positive balancd Account Lebvel

                    int amount = 0;
                    string WithdrawlLimit = "";

                    Paragraph paragraph13 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph13.Range.Text = "\nS12:6-Account Positive balance\n";
                    paragraph13.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph13.Range.InsertParagraphAfter();
                    Paragraph paragraph14 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph14.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range7 = paragraph14.Range;
                    range7.Text = "Contract Type\tWithdrawl Limit\n";
                    List_of_dictionaries.c.CommandText = "select name,sum(amount),sum(prc),decode(sum(type),1,'As Total Of',2,'As Maximum Of',3,'As Minimum Of','Not Used (100%)')  Type,c currency from (select t2.NAME name,t1.VALUE amount,null prc,'International' c,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile  where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'ACCASHLIMITPBTYPEINT'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch  and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,null prc,'International' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'ACCASHLIMITPBTYPEINT'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'International' c,null Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'ACCASHLIMITPBTYPEINT' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,t1.VALUE amount,null prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'ACCCASHLIMITPBTYPEDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID  and t2.status like '1' union all select t2.NAME name,null amount,null prc,'Domestic' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile  where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'ACCCASHLIMITPBTYPEDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'ACCCASHLIMITPBTYPEDOM' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID) group by name ,c order by 1";

                    OracleDataReader oracleDataReader6 = List_of_dictionaries.c.ExecuteReader();

                    while (oracleDataReader6.Read())
                    {

                        if (oracleDataReader6[1].ToString() != null || oracleDataReader6[1].ToString() != "")
                        {
                            amount = Convert.ToInt32("0" + oracleDataReader6[1].ToString());
                            switch (amount)
                            {
                                case 1:
                                    WithdrawlLimit = "Calculate From Total Limit";
                                    break;
                                case 2:
                                    WithdrawlLimit = "Calculate From Credit Limit";
                                    break;
                                case 3:
                                    WithdrawlLimit = "Increase By Positive Balance";
                                    break;
                                default:
                                    WithdrawlLimit = "Not Defined";
                                    break;
                            }
                        }
                        Microsoft.Office.Interop.Word.Range range2 = range7;
                        string str = range2.Text + oracleDataReader6[0] + "\t" + WithdrawlLimit + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader6.Close();
                    range7.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range7.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range7.Tables[1].Borders.Enable = 1;
                    // end of Account

                    //Positive balancd Primary Lebvel
                    Paragraph paragraph15 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph15.Range.Text = "\nS12:7-Primary Card Positive balance\n";
                    paragraph15.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph15.Range.InsertParagraphAfter();
                    Paragraph paragraph16 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph16.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range8 = paragraph16.Range;
                    range8.Text = "Contract Type\tWithdrawl Limit\n";
                    List_of_dictionaries.c.CommandText = "select name,sum(amount),sum(prc),decode(sum(type),1,'As Total Of',2,'As Maximum Of',3,'As Minimum Of','Not Used (100%)')  Type,c currency from (select t2.NAME name,t1.VALUE amount,null prc,'International' c,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile  where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICASHLIMITPBTYPEINT'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch  and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,null prc,'International' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICASHLIMITPBTYPEINT'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'International' c,null Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICASHLIMITPBTYPEINT' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,t1.VALUE amount,null prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICASHLIMITPBTYPEDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID  and t2.status like '1' union all select t2.NAME name,null amount,null prc,'Domestic' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile  where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICASHLIMITPBTYPEDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'PRICASHLIMITPBTYPEDOM' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID) group by name ,c order by 1";

                    OracleDataReader oracleDataReader7 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader7.Read())
                    {
                        if (oracleDataReader7[1].ToString() != null || oracleDataReader7[1].ToString() != "")
                        {
                            amount = Convert.ToInt32("0" + oracleDataReader7[1].ToString());
                            switch (amount)
                            {
                                case 1:
                                    WithdrawlLimit = "Calculate From Total Limit";
                                    break;
                                case 2:
                                    WithdrawlLimit = "Calculate From Credit Limit";
                                    break;
                                case 3:
                                    WithdrawlLimit = "Increase By Positive Balance";
                                    break;
                                default:
                                    WithdrawlLimit = "Not Defined";
                                    break;
                            }
                        }
                        Microsoft.Office.Interop.Word.Range range2 = range8;
                        string str = range2.Text + oracleDataReader7[0] + "\t" + WithdrawlLimit + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader7.Close();
                    range8.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range8.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range8.Tables[1].Borders.Enable = 1;
                    // end of Primary

                    //Positive balance supplementary  Lebvel
                    Paragraph paragraph17 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph17.Range.Text = "\nS12:8-Supplementary Card Positive balance\n";
                    paragraph17.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph17.Range.InsertParagraphAfter();
                    Paragraph paragraph18 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph18.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range9 = paragraph18.Range;
                    range9.Text = "Contract Type\tWithdrawlLimit\n";
                    List_of_dictionaries.c.CommandText = "select name,sum(amount),sum(prc),decode(sum(type),1,'As Total Of',2,'As Maximum Of',3,'As Minimum Of','Not Used (100%)')  Type,c currency from (select t2.NAME name,t1.VALUE amount,null prc,'International' c,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile  where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCASHLIMITPBTYPEINT'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch  and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,null prc,'International' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + " and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCASHLIMITPBTYPEINT'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'International' c,null Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCASHLIMITPBTYPEINT' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEINT' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,t1.VALUE amount,null prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCASHLIMITPBTYPEDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID  and t2.status like '1' union all select t2.NAME name,null amount,null prc,'Domestic' c ,t1.VALUE Type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile  where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCASHLIMITPBTYPEDOM'and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID and t2.status like '1' union all select t2.NAME name,null amount,t1.VALUE prc,'Domestic' c ,null type from a4m.tcontracttypeparameters t1,a4m.tcontracttype t2,a4m.tcontracttypeparameters tt1,a4m.tcontractprofile where t1.branch=" + Frm_1.bank_num + "and t1.BRANCH=t2.BRANCH and t1.CONTRACTTYPE=t2.TYPE and upper(t1.KEY)like'SUPCASHLIMITPBTYPEDOM' and tt1.branch=t1.branch and tt1.CONTRACTTYPE=t1.CONTRACTTYPE and upper(tt1.KEY) Like 'PROFILEDOM' and tt1.branch=tcontractprofile.branch and tt1.VALUE=tcontractprofile.PROFILEID) group by name ,c order by 1";

                    OracleDataReader oracleDataReader8 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader8.Read())
                    {
                        if (oracleDataReader8[1].ToString() != null || oracleDataReader8[1].ToString() != "")
                        {
                            amount = Convert.ToInt32("0" + oracleDataReader8[1].ToString());
                            switch (amount)
                            {
                                case 1:
                                    WithdrawlLimit = "Calculate From Total Limit";
                                    break;
                                case 2:
                                    WithdrawlLimit = "Calculate From Credit Limit";
                                    break;
                                case 3:
                                    WithdrawlLimit = "Increase By Positive Balance";
                                    break;
                                default:
                                    WithdrawlLimit = "Not Defined";
                                    break;
                            }
                        }
                        Microsoft.Office.Interop.Word.Range range2 = range9;
                        string str = range2.Text + oracleDataReader8[0] + "\t" + WithdrawlLimit + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader8.Close();
                    range9.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range9.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range9.Tables[1].Borders.Enable = 1;
                    // end of Supp.

                    //iatta


                    oPara1 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                }
                if (flag13)
                {
                    oPara1.Range.Text = "\fInstallment Setting 13 \n";
                    oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara1.Range.Font.Bold = 0;
                    oPara1.Range.Font.Size = 9f;
                    oPara1.Format.SpaceAfter = 0.0f;
                    oPara1.Range.InsertParagraphAfter();

                    Paragraph paragraph1 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph1.Range.Text = "\nBeginning Of Installment\n";
                    paragraph1.Range.InsertParagraphAfter();
                    paragraph1.Range.Font.Underline = WdUnderline.wdUnderlineNone;

                    OracleCommand command1 = Frm_1.dbcon.CreateCommand();
                    command1.CommandText = BaseInstallment;
                    OracleDataReader oracleDataReader1 = command1.ExecuteReader();

                    while (oracleDataReader1.Read())
                    {
                        ContractNumber = oracleDataReader1[1].ToString();
                        string installmentName = oracleDataReader1[0].ToString();

                        // === INSTALLMENT NAME HEADER ===
                        Paragraph installmentHeader = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        installmentHeader.Range.Font.Bold = 1;
                        installmentHeader.Range.Font.Size = 10f;
                        installmentHeader.Range.Text = "\nInstallment: " + installmentName + "\n";
                        installmentHeader.Range.InsertParagraphAfter();
                        installmentHeader.Range.Font.Bold = 0;

                        // === TABLE 1: Name/Value Parameters ===
                        Microsoft.Office.Interop.Word.Range range3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt).Range;
                        range3.Text = "Name\tValue\n";

                        OracleCommand command2 = Frm_1.dbcon.CreateCommand();
                        command2.CommandText = $"SELECT KEY, DECODE(KEY, 'DAYSINYEAR', DECODE(VALUE, '1', 'Native','2', '360/30', VALUE), 'COUNTDAYSMODE', DECODE(VALUE, '1', 'days in cycle','2', 'days in month', VALUE), 'CALCMETHOD', DECODE(VALUE, '1', 'Initial balance, fixed repayment','2', 'Unpaid balance, fixed repayment','3','Unpaid balance, decreasing repayment','4','Unpaid balance, fixed repayment per month', VALUE), 'CALCSCHEDULEBY', DECODE(VALUE, '1', 'Count of billing cycles','2', 'Amount of regular repayment', VALUE), 'INSTALLMENTCT', DECODE(VALUE, '1', 'Charge interest','2', 'As fee', '3', 'As prorated fee',VALUE), 'BASE2ACCELERATION', DECODE(VALUE, '1', 'Dont accrue','2', 'Accrue', VALUE), 'PENALTYCHARGETYPE', DECODE(VALUE, '1', 'Do not charge','2', 'Charge as interest','3','Charge as fee', VALUE), 'LIMITMETHOD', DECODE(VALUE, '1', 'As total of','2', 'As maximum between','3','As minimum between', VALUE), 'ACCELDATEID', DECODE(VALUE, '1', 'Current Date','2', 'Statement Date', VALUE), 'MULTIPLELOAN', DECODE(VALUE, '1', 'Allowed','2', 'Not allowed', VALUE), 'FEEMODE', DECODE(VALUE, '1', 'Do not charge','2', 'From RCM settings','3','From Installment settings','4','RCM and Installment', VALUE), 'FEEAMOUNT', DECODE(VALUE, '1', 'Remaining debt','2', 'Full amount','3','Unpaid principal amount', VALUE), 'FIXEDINTEREST', DECODE(VALUE, '1', 'New tranches only','2', 'All tranches', VALUE), VALUE) AS VALUE, D.CONTRACTTYPE FROM A4m.TCONTRACTTYPEPARAMETERS d where branch = {Frm_1.bank_num} and CONTRACTTYPE = {ContractNumber}";
                        OracleDataReader oracleDataReader2 = command2.ExecuteReader();

                        while (oracleDataReader2.Read())
                        {
                            range3.InsertAfter(oracleDataReader2[0].ToString() + "\t" + oracleDataReader2[1].ToString() + "\n");
                        }
                        oracleDataReader2.Close();
                        command2.Dispose();

                        List_of_dictionaries.col_width = 120;
                        range3.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                        range3.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                        range3.Tables[1].Borders.Enable = 1;

                        // === TABLE 2: Linked Contracts ===
                        Paragraph linkedHeader = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        linkedHeader.Range.Text = "\nLinked Contracts:\n";
                        linkedHeader.Range.InsertParagraphAfter();

                        Microsoft.Office.Interop.Word.Range range5 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt).Range;
                        range5.Text = "Contract Name\n";

                        OracleCommand command3 = Frm_1.dbcon.CreateCommand();
                        command3.CommandText = $"select CONTRACT.NAME from A4M.TCONTRACTTYPE contract join A4M.TCONTRACTTYPELink link on LINK.MAINTYPE = CONTRACT.TYPE where CONTRACT.BRANCH = {Frm_1.bank_num} and SCHEMATYPE != 3 and LINKTYPE = {ContractNumber}";
                        OracleDataReader oracleDataReader3 = command3.ExecuteReader();

                        bool hasLinkedContracts = false;
                        while (oracleDataReader3.Read())
                        {
                            hasLinkedContracts = true;
                            range5.InsertAfter(oracleDataReader3[0].ToString() + "\n");
                        }
                        oracleDataReader3.Close();
                        command3.Dispose();

                        if (hasLinkedContracts)
                        {
                            List_of_dictionaries.col_width = 120;
                            range5.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                            range5.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                            range5.Tables[1].Borders.Enable = 1;
                        }

                        // Add spacing between installments
                        Paragraph spacer = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                        spacer.Range.Text = "\n";
                        spacer.Range.InsertParagraphAfter();
                    }
                    oracleDataReader1.Close();

                }
                if (flag14)
                {
                    oPara1.Range.Text = "\fSection 13 : Appendix\n";
                    oPara1.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    oPara1.Range.Font.Bold = 0;
                    oPara1.Range.Font.Size = 9f;
                    oPara1.Format.SpaceAfter = 0.0f;
                    oPara1.Range.InsertParagraphAfter();
                    Paragraph paragraph2 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph2.Range.Text = "Card Status\n";
                    paragraph2.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph2.Range.InsertParagraphAfter();
                    Paragraph paragraph3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph3.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range1 = paragraph3.Range;
                    range1.Text = "Card Status Code\tCard Status Name\n";
                    List_of_dictionaries.c.CommandText = "select treferencecrd_stat.crd_stat code,name from a4m.treferencecrd_stat order by 1";
                    OracleDataReader oracleDataReader1 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader1.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range1;
                        string str = range2.Text + oracleDataReader1[0] + "\t" + oracleDataReader1[1] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader1.Close();
                    range1.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range1.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range1.Tables[1].Borders.Enable = 1;
                    Paragraph paragraph4 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph4.Range.Text = "\nCard State\n";
                    paragraph4.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    paragraph4.Range.InsertParagraphAfter();
                    Paragraph paragraph5 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                    paragraph5.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    Microsoft.Office.Interop.Word.Range range3 = paragraph5.Range;
                    range3.Text = "Card State Code\tCard State Name\n";
                    List_of_dictionaries.c.CommandText = "select treferencecardsign.cardsign code,treferencecardsign.name from a4m.treferencecardsign where branch=" + Frm_1.bank_num + " order by 1";
                    OracleDataReader oracleDataReader2 = List_of_dictionaries.c.ExecuteReader();
                    while (oracleDataReader2.Read())
                    {
                        Microsoft.Office.Interop.Word.Range range2 = range3;
                        string str = range2.Text + oracleDataReader2[0] + "\t" + oracleDataReader2[1] + "\n";
                        range2.Text = str;
                    }
                    oracleDataReader2.Close();
                    range3.ConvertToTable(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.col_width, ref List_of_dictionaries.format, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.apply_color, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.auto_fit_true, ref List_of_dictionaries.auto_fit, ref List_of_dictionaries.m_objOpt);
                    range3.Tables[1].TableDirection = WdTableDirection.wdTableDirectionLtr;
                    range3.Tables[1].Borders.Enable = 1;
                }
                List_of_dictionaries.oDoc.Protect(WdProtectionType.wdAllowOnlyReading, false, "Crystal_2014", List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt);
                //List_of_dictionaries.oDoc.SaveAs(ref List_of_dictionaries.filename, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                List_of_dictionaries.oDoc.SaveAs(ref List_of_dictionaries.filename, WdSaveFormat.wdFormatDocumentDefault, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                List_of_dictionaries.oDoc.Close(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                //List_of_dictionaries.oWord.Quit(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                //EDT-975
                ((Microsoft.Office.Interop.Word._Application)List_of_dictionaries.oWord).Quit(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                GC.Collect();
                int num = (int)MessageBox.Show("Report Successfully Generated", "Message", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            catch (System.Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (!((this.saveFileDialog1.ShowDialog()).ToString().ToLower() == "ok"))
                    return;
                List_of_dictionaries.filename = this.saveFileDialog1.FileName;
                List_of_dictionaries.oWord = (Word.Application)new Word.Application();
                List_of_dictionaries.oDoc = (_Document)List_of_dictionaries.oWord.Documents.Add(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                List_of_dictionaries.oDoc.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                List_of_dictionaries.oDoc.PageSetup.RightMargin = 20f;
                List_of_dictionaries.oDoc.PageSetup.LeftMargin = 20f;
                List_of_dictionaries.oDoc.PageSetup.TopMargin = 20f;
                List_of_dictionaries.oDoc.PageSetup.BottomMargin = 20f;
                List_of_dictionaries.oDoc.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
                List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                try
                {
                    List_of_dictionaries.oWord.Selection.InlineShapes.AddPicture("C:\\Compass\\Operation Software\\MSCCLOGO[1].gif", ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                }
                catch
                {
                }
                Paragraph paragraph1 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                paragraph1.Range.Text = "Service Code Form";
                paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph1.Range.Font.Underline = WdUnderline.wdUnderlineDouble;
                paragraph1.Range.Font.Name = "Verdana";
                paragraph1.Range.Font.Color = WdColor.wdColorAutomatic;
                paragraph1.Range.Font.Size = 15f;
                paragraph1.Format.SpaceAfter = 18f;
                paragraph1.Range.InsertParagraphAfter();
                Paragraph paragraph2 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                paragraph2.Range.Text = " ";
                paragraph2.Range.Font.Size = 10f;
                paragraph2.Format.SpaceAfter = 3f;
                paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                Paragraph paragraph3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                // ISSUE: explicit reference operation
                Microsoft.Office.Interop.Word.Range range1 = List_of_dictionaries.oDoc.Bookmarks[@List_of_dictionaries.oEndOfDoc].Range;
                Word.Table table1 = List_of_dictionaries.oDoc.Tables.Add(range1, 3, 2, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                table1.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                table1.AllowAutoFit = true;
                table1.Cell(1, 1).Range.Text = "Bank Name";
                table1.Cell(1, 1).Range.Font.Color = WdColor.wdColorAutomatic;
                table1.Cell(1, 1).Range.Font.Bold = 1;
                table1.Cell(1, 1).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                table1.Cell(1, 2).Range.Text = Frm_1.bank_name;
                table1.Cell(1, 2).Range.Font.Bold = 0;
                table1.Cell(1, 2).Range.Font.Color = WdColor.wdColorAutomatic;
                table1.Cell(1, 2).Range.Font.Underline = WdUnderline.wdUnderlineNone;
                table1.Cell(2, 1).Range.Text = "Business ID";
                table1.Cell(2, 1).Range.Font.Color = WdColor.wdColorAutomatic;
                table1.Cell(2, 1).Range.Font.Bold = 1;
                table1.Cell(2, 1).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                table1.Cell(2, 2).Range.Text = Frm_1.Business_ID;
                table1.Cell(2, 2).Range.Font.Bold = 0;
                table1.Cell(2, 2).Range.Font.Color = WdColor.wdColorAutomatic;
                table1.Cell(2, 2).Range.Font.Underline = WdUnderline.wdUnderlineNone;
                table1.Cell(3, 1).Range.Text = "Country Code";
                table1.Cell(3, 1).Range.Font.Color = WdColor.wdColorAutomatic;
                table1.Cell(3, 1).Range.Font.Bold = 1;
                table1.Cell(3, 1).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                table1.Cell(3, 2).Range.Text = Frm_1.country_code;
                table1.Cell(3, 2).Range.Font.Bold = 0;
                table1.Cell(3, 2).Range.Font.Color = WdColor.wdColorAutomatic;
                table1.Cell(3, 2).Range.Font.Underline = WdUnderline.wdUnderlineNone;
                paragraph3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                OracleCommand command = Frm_1.dbcon.CreateCommand();
                command.CommandText = "select treferencecardproduct.name,substr(prefix,1,6),period,servicecode,pvki,treferencecardsign.NAME,treferencecrd_stat.NAME from a4m.treferencecardproduct,a4m.treferencecardsign,a4m.treferencecrd_stat where treferencecardproduct.branch=treferencecardsign.branch(+) and treferencecardproduct.STATE=treferencecardsign.CARDSIGN(+) and treferencecardproduct.STATUS=treferencecrd_stat.CRD_STAT(+) and treferencecardproduct.branch=" + Frm_1.bank_num + " order by 1";
                OracleDataReader oracleDataReader1 = command.ExecuteReader();
                int NumRows = 1;
                while (oracleDataReader1.Read())
                    ++NumRows;
                oracleDataReader1.Close();
                // ISSUE: explicit reference operation
                Microsoft.Office.Interop.Word.Range range2 = List_of_dictionaries.oDoc.Bookmarks[@List_of_dictionaries.oEndOfDoc].Range;
                Word.Table table2 = List_of_dictionaries.oDoc.Tables.Add(range2, NumRows, 5, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                table2.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                table2.LeftPadding = 1f;
                table2.Cell(1, 1).Range.Text = "Product";
                table2.Cell(1, 1).Range.Font.Color = WdColor.wdColorAutomatic;
                table2.Cell(1, 1).Range.Font.Bold = 1;
                table2.Cell(1, 1).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                table2.Cell(1, 2).Range.Text = "BIN";
                table2.Cell(1, 2).Range.Font.Color = WdColor.wdColorAutomatic;
                table2.Cell(1, 2).Range.Font.Bold = 1;
                table2.Cell(1, 2).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                table2.Cell(1, 3).Range.Text = "Service Code";
                table2.Cell(1, 3).Range.Font.Color = WdColor.wdColorAutomatic;
                table2.Cell(1, 3).Range.Font.Bold = 1;
                table2.Cell(1, 3).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                table2.Cell(1, 4).Range.Text = "Card Manufacture";
                table2.Cell(1, 4).Range.Font.Color = WdColor.wdColorAutomatic;
                table2.Cell(1, 4).Range.Font.Bold = 1;
                table2.Cell(1, 4).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                table2.Cell(1, 5).Range.Text = "Product Type";
                table2.Cell(1, 5).Range.Font.Color = WdColor.wdColorAutomatic;
                table2.Cell(1, 5).Range.Font.Bold = 1;
                table2.Cell(1, 5).Range.Font.Underline = WdUnderline.wdUnderlineSingle;
                command.CommandText = "select treferencecardproduct.name,substr(prefix,1,6),period,servicecode,pvki,treferencecardsign.NAME,treferencecrd_stat.NAME from a4m.treferencecardproduct,a4m.treferencecardsign,a4m.treferencecrd_stat where treferencecardproduct.branch=treferencecardsign.branch(+) and treferencecardproduct.STATE=treferencecardsign.CARDSIGN(+) and treferencecardproduct.STATUS=treferencecrd_stat.CRD_STAT(+) and treferencecardproduct.branch=" + Frm_1.bank_num + " order by 1";
                OracleDataReader oracleDataReader2 = command.ExecuteReader();
                int Row = 2;
                while (oracleDataReader2.Read())
                {
                    table2.Cell(Row, 1).Range.Text = oracleDataReader2[0].ToString();
                    table2.Cell(Row, 1).Range.Font.Bold = 0;
                    table2.Cell(Row, 1).Range.Font.Color = WdColor.wdColorAutomatic;
                    table2.Cell(Row, 1).Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    table2.Cell(Row, 2).Range.Text = oracleDataReader2[1].ToString();
                    table2.Cell(Row, 2).Range.Font.Bold = 0;
                    table2.Cell(Row, 2).Range.Font.Color = WdColor.wdColorAutomatic;
                    table2.Cell(Row, 2).Range.Font.Underline = WdUnderline.wdUnderlineNone;
                    ++Row;
                }
                oracleDataReader2.Close();
                Paragraph paragraph4 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                paragraph4.Range.Text = "\n\nAuthorized Signature";
                paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph4.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                paragraph4.Range.Font.Color = WdColor.wdColorAutomatic;
                paragraph4.Range.Font.Size = 12f;
                paragraph4.Range.Font.Bold = 1;
                paragraph3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                List_of_dictionaries.oDoc.Protect(WdProtectionType.wdAllowOnlyReading, false, "Crystal_2014", List_of_dictionaries.m_objOpt, List_of_dictionaries.m_objOpt);
                List_of_dictionaries.oDoc.SaveAs(ref List_of_dictionaries.filename, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                List_of_dictionaries.oDoc.Close(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                int num = (int)MessageBox.Show("Service Code Form Successfully Generated", "Message", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            catch (System.Exception ex)
            {
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                if (!((this.saveFileDialog1.ShowDialog()).ToString().ToLower() == "ok"))
                    return;
                List_of_dictionaries.filename = this.saveFileDialog1.FileName;
                List_of_dictionaries.oWord = (Word.Application)new Word.Application();
                List_of_dictionaries.oDoc = (_Document)List_of_dictionaries.oWord.Documents.Add(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                List_of_dictionaries.oDoc.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                List_of_dictionaries.oDoc.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
                List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                Paragraph paragraph1 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                paragraph1.Range.Text = "Service Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nService Code Form\nv";
                paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph1.Range.Font.Underline = WdUnderline.wdUnderlineDouble;
                paragraph1.Range.Font.Name = "Verdana";
                paragraph1.Range.Font.Color = WdColor.wdColorAutomatic;
                paragraph1.Range.Font.Size = 15f;
                paragraph1.Format.SpaceAfter = 18f;
                paragraph1.Range.InsertParagraphAfter();
                Paragraph paragraph2 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                paragraph2.Range.Text = " Hello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\nHello\n";
                paragraph2.Range.Font.Size = 10f;
                paragraph2.Format.SpaceAfter = 3f;
                paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                Paragraph paragraph3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                Paragraph paragraph4 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                paragraph4.Range.Text = "\n\nAuthorized Signature";
                paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph4.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                paragraph4.Range.Font.Color = WdColor.wdColorAutomatic;
                paragraph4.Range.Font.Size = 12f;
                paragraph4.Range.Font.Bold = 1;
                Paragraph paragraph5 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                paragraph5.Range.Text = "\n\nAuthorized Signature";
                paragraph5.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph5.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                paragraph5.Range.Font.Color = WdColor.wdColorAutomatic;
                paragraph5.Range.Font.Size = 12f;
                paragraph5.Range.Font.Bold = 1;
                paragraph3 = List_of_dictionaries.oDoc.Content.Paragraphs.Add(ref List_of_dictionaries.m_objOpt);
                List_of_dictionaries.oDoc.SaveAs(ref List_of_dictionaries.filename, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                List_of_dictionaries.oDoc.Close(ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt, ref List_of_dictionaries.m_objOpt);
                int num = (int)MessageBox.Show("Service Code Form Successfully Generated", "Message", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            catch (System.Exception ex)
            {
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.linkLabel2.Text == "Select All")
            {
                for (int index = 0; index < this.checkedListBox2.Items.Count; ++index)
                    this.checkedListBox2.SetItemChecked(index, true);
                this.linkLabel2.Text = "DeSelect All";
            }
            else
            {
                if (!(this.linkLabel2.Text == "DeSelect All"))
                    return;
                for (int index = 0; index < this.checkedListBox2.Items.Count; ++index)
                    this.checkedListBox2.SetItemChecked(index, false);
                this.linkLabel2.Text = "Select All";
            }
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}