package mobilehome.info;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class MobileHomeInfo {
	
	public final static int CT = 1;
	public final static int MESA = 2;
	private float ExpectedMonthlyRent;
	private float PreviousBalance;
	private float LateFee;
	private float Credit;
	private float ReceivedBefore5th;
	private int LotNumber;
	public final static int CT_MAX_LOTS=27; //25 is max but excel has header row
	public final static int MESA_MAX_LOTS=31; //25 is max but excel has header row
	public final static int LATEFEE_BEFORE_5TH = 50; //late fee before 5th of the month
	
	public int getLotNumber(){return LotNumber;}
	public float getCredit(){return Credit;}
	public float getLateFee(){return LateFee;}
	public float getPreviousBalance(){return PreviousBalance;}
	public float getExpectedMonthlyRent(){return ExpectedMonthlyRent;}
	public float getReceivedBefore5th(){return ReceivedBefore5th;}

	public final static int LISTALLHOMES=100;
	
	public static String LotBalanceURL(int n){
		int column = 17; //Column 'Q'
		int minrow = 0;
		int maxrow = 0;
		if(n == 100){
			minrow = 2;
			maxrow = 26;
		}else{
			minrow=1+n;
			maxrow=minrow;
		}
		return "?min-row="+minrow+"&min-col="+column+"&max-row="+maxrow+"&max-col="+column;
	}
	
	public MobileHomeInfo(float expectedmonthlyrent, 
						float previousbalance,
						float latefee,
						float credit,
						int lotnumber,
						float receivedbefore5th){
		ExpectedMonthlyRent = expectedmonthlyrent;
		PreviousBalance = previousbalance;
		LateFee = latefee;
		Credit = credit;
		LotNumber = lotnumber;
		ReceivedBefore5th = receivedbefore5th;
	}
	
	public String toString(){
		return "LotNum: " + LotNumber + ": ExpectedMonthlyRent: " + ExpectedMonthlyRent + ", PreviousBalance: " + PreviousBalance
				+ ", LateFee: " + LateFee
				+ ", Credit: " + Credit;
	}
	
	public float TotalDue(){
		return ExpectedMonthlyRent + PreviousBalance + LateFee + Credit;
	}
	
	public static void GenerateLateNotices(List<MobileHomeInfo> mobilehomeinfo, int mobilepark) 
			throws BiffException, IOException, RowsExceededException, WriteException, ParseException{
		FileOutputStream fout = null;
		String park = null;
		String address = null;
		String city_zip = "Palestine, Texas 75801";
		String mgrcontact = "Park Manager Phone: (903)600­0647";
		String email = null;
		Date dt = new Date();
		SimpleDateFormat filenameformat = new SimpleDateFormat("yyyymmdd");
		
		if(mobilepark == MobileHomeInfo.CT){
			fout = new FileOutputStream("resources/" + filenameformat.format(dt) + "_CT-LateNotice.docx");
			park = "Cross Timbers Mobile Home Park";
			address = "4507 West Oak Street";
			email = "crosstimbersmhp@yahoo.com";
		}
		else{
			fout = new FileOutputStream("resources/" + filenameformat.format(dt) + "_Mesa-LateNotice.docx");
			park = "Mesa Mobile Home Park";
			address = "1118 North Fort Street";
			email = "mesamhp@yahoo.com";
}
		
		XWPFDocument doc = new XWPFDocument();

		for(MobileHomeInfo mh : mobilehomeinfo){
			float amountdue = mh.getExpectedMonthlyRent()+mh.getLateFee()+mh.getPreviousBalance()+mh.getCredit();
			if(amountdue > mh.getReceivedBefore5th()){
				float amountafter5th  = amountdue-mh.getReceivedBefore5th()+50;
				float amountafter10th = amountdue-mh.getReceivedBefore5th()+100;

				//Line 1
				XWPFParagraph p1 = doc.createParagraph();
				p1.setPageBreak(true);
				p1.setAlignment(ParagraphAlignment.CENTER);
				XWPFRun p1run = p1.createRun();
				p1run.setFontFamily("Times New Roman");
				p1run.setText(park);
				//Line 2
				XWPFParagraph p2 = doc.createParagraph();
				p2.setAlignment(ParagraphAlignment.CENTER);
				XWPFRun p2run = p2.createRun();
				p2run.setFontFamily("Times New Roman");
				p2run.setText(address);
				//Line 3
				XWPFParagraph p3 = doc.createParagraph();
				p3.setAlignment(ParagraphAlignment.CENTER);
				XWPFRun p3run = p3.createRun();
				p3run.setFontFamily("Times New Roman");
				p3run.setText(city_zip);
				//Line 4
				XWPFParagraph p4 = doc.createParagraph();
				p4.setAlignment(ParagraphAlignment.CENTER);
				XWPFRun p4run = p4.createRun();
				p4run.setFontFamily("Times New Roman");
				p4run.setText(mgrcontact);
				//Line 5
				XWPFParagraph p5 = doc.createParagraph();
				p5.setAlignment(ParagraphAlignment.CENTER);
				XWPFRun p5run = p5.createRun();
				p5run.setFontFamily("Times New Roman");
				p5run.setText(email);
				//Line 6
				XWPFParagraph p6 = doc.createParagraph();
				p6.setAlignment(ParagraphAlignment.CENTER);
				XWPFRun p6run = p6.createRun();
				p6run.setFontFamily("Times New Roman");
				p6run.setText("");
				//Line 7
				XWPFParagraph p7 = doc.createParagraph();
				p7.setAlignment(ParagraphAlignment.CENTER);
				XWPFRun p7run = p7.createRun();
				p7run.setFontFamily("Times New Roman");
				p7run.setText("");
				//Line 8
				XWPFParagraph p8 = doc.createParagraph();
				p8.setAlignment(ParagraphAlignment.LEFT);
				XWPFRun p8run = p8.createRun();
			    SimpleDateFormat sdf = new SimpleDateFormat("dd MMMM yyyy");
			    Date now = new Date();
			    p8run.setFontFamily("Times New Roman");
				p8run.setText("Date: " + sdf.format(now));
				//Line 9
				XWPFRun p9run = doc.createParagraph().createRun();
				p9run.setText("");
				//Line 10
				XWPFRun p10run = doc.createParagraph().createRun();
				p10run.setFontFamily("Times New Roman");
				p10run.setText("Lot# " + mh.getLotNumber());
				//Line 11
				XWPFRun p11run = doc.createParagraph().createRun();
				p11run.setText("");
				//Line 12
				XWPFRun p12run = doc.createParagraph().createRun();
				p12run.setFontFamily("Times New Roman");
				p12run.setText("You have an unpaid balance of: " + amountafter5th + "$");
				//Line 13
				doc.createParagraph().createRun().setText("");
				//Line 14
				XWPFRun p14run = doc.createParagraph().createRun();
				p14run.setFontFamily("Times New Roman");
				Calendar c = Calendar.getInstance();
				c.add(Calendar.DATE, 5);
				SimpleDateFormat sdf2 = new SimpleDateFormat("dd MMMM yyyy");
				p14run.setText("This letter will be your final warning before you are turned over " + 
				               "to the attorney’s office in 36 hours for an eviction to be filed. " + 
						       "You were given a notice to pay the amount due in full. Rent is due on " + 
				               "the 1st and considered late if paid after the 5th of each month. Rent " + 
						       "amount due after " + sdf2.format(c.getTime()) + " will be " + amountafter10th);
				//Line 15
				doc.createParagraph().createRun().setText("");
				//Line 16
				XWPFRun p16run = doc.createParagraph().createRun();
				p16run.setFontFamily("Times New Roman");
				p16run.setText("We have given all tenants a grace period of 5 days to pay the amount owed " + 
				               "in full and expect tenants not to take advantage of our policy.  We will " + 
						       "begin evicting those that are continually receiving this letter for nonpayment " + 
				               "and paying late. I am willing to discuss your account and take full payment only. "); 

				//Line 17
				doc.createParagraph().createRun().setText("");
				//Line 18
				doc.createParagraph().createRun().setText("");
				//Line 19
				doc.createParagraph().createRun().setText("");
				//Line 20
				XWPFRun p20run = doc.createParagraph().createRun();
				p20run.setFontFamily("Times New Roman");
				p20run.setBold(true);
				p20run.setText(park + " Management");
	
			}
		}//for
	    doc.write(fout);
	    fout.close();
	    System.out.println("Job completed");
	}
}
