import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.JFileChooser;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class NewJFrame extends javax.swing.JFrame {

	File exFile;
	File inFile;
	int nuberOfCells;
	ArrayList<String> alMSISDN = new ArrayList<String>();
    ArrayList<String> alICCID = new ArrayList<String>();
    String StringtotNoOfOrders;
    StringBuffer finalStringOfXML = new StringBuffer();
    String s;
    int totNoOfOrders;
    String swe;
    /**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	/**
     * Creates new form NewJFrame
     */
    public NewJFrame() {
        initComponents();
    }

    
    // <editor-fold defaultstate="collapsed" desc="Generated Code">                          
    private void initComponents() {

        jScrollPane1 = new javax.swing.JScrollPane();
        jTextArea1 = new javax.swing.JTextArea();
        jButton1 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        jTextField1 = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();
        jButton3 = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jTextArea1.setColumns(20);
        jTextArea1.setRows(5);
        jScrollPane1.setViewportView(jTextArea1);

        jButton1.setText("Select Input logs File");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        jButton2.setText("Select Data Excel File");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        jLabel1.setText("No. Of Orders");

        jButton3.setText("Submit");
        jButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton3ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addGroup(layout.createSequentialGroup()
                        .addContainerGap()
                        .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 934, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(28, 28, 28)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 196, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(60, 60, 60)
                                .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 197, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 141, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(38, 38, 38)
                                .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, 78, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(jButton3, javax.swing.GroupLayout.PREFERRED_SIZE, 103, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(66, 66, 66)))))
                .addContainerGap(23, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(24, 24, 24)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jButton1, javax.swing.GroupLayout.DEFAULT_SIZE, 76, Short.MAX_VALUE)
                    .addComponent(jButton2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addGap(28, 28, 28)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, 54, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 54, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton3, javax.swing.GroupLayout.PREFERRED_SIZE, 41, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(34, 34, 34)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 552, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(30, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>                        

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {
    	this.setInputFile();

    }                                        

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {                                         
    	try {
        	this.loadICCIDandMSISDN();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
        }
    }                                        

    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {                                         
    	try {
    		setNoOfOrders();
            swe = new String(this.everyLogic());
            jTextArea1.append(swe);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }      
//-------------------------------------------------------------------------------------------------    
//-------------------------------------------------------------------    
    public void setNoOfOrders(){
    	s = jTextField1.getText();
    	System.out.println(s);
    }

//---------------------------------------------------------
    
    
    File inputFile;
    public File setInputFile(){
    	JFileChooser fileChooser = new JFileChooser();
       int returnVal = fileChooser.showOpenDialog(null);
       if(returnVal == JFileChooser.APPROVE_OPTION){
    	   inputFile = fileChooser.getSelectedFile();
       }
       return inputFile;
    }

//---------------------------------------------------------------------------------------------    
    
    File excelFile;
    public File retexcelFile(){
    	JFileChooser fileChooser = new JFileChooser();
       int returnVal = fileChooser.showOpenDialog(null);
       if(returnVal == JFileChooser.APPROVE_OPTION){
    	   excelFile = fileChooser.getSelectedFile();
       }
       return excelFile;
    }
//--------------------------------------------------------------------------------------------
    
    
    public void loadICCIDandMSISDN() throws IOException{
   	    exFile = this.retexcelFile();
   	 	FileInputStream excelFile = new FileInputStream(exFile);
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet datatypeSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = datatypeSheet.iterator();


       
   	 nuberOfCells = -1 ;
        
       	 while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (cellIterator.hasNext()) {
               	 nuberOfCells++;

                    Cell currentCell = cellIterator.next();
                    if (currentCell.getCellTypeEnum() == CellType.STRING && nuberOfCells == 0) {
                        alMSISDN.add(currentCell.getStringCellValue());
                        }
                        else if (currentCell.getCellTypeEnum() == CellType.STRING && nuberOfCells%2 == 0) {
                        alMSISDN.add(currentCell.getStringCellValue());
                    }
                    else if (currentCell.getCellTypeEnum() == CellType.STRING && nuberOfCells%2 != 0) {
                   	 alICCID.add(currentCell.getStringCellValue());
                    }
                }

            }
            workbook.close(); 
   	}
    
//-------------------------------------------------------------------------------------------------------------
    
  //-------------------------------------------------------------------------------
    public StringBuffer everyLogic() throws IOException {
      	 StringBuffer Tex;
      	StringtotNoOfOrders = this.s;
      	totNoOfOrders = Integer.parseInt(StringtotNoOfOrders); 
   	    System.out.println(StringtotNoOfOrders);

         
         try {
       	 @SuppressWarnings("deprecation")
   		String testHtml = FileUtils.readFileToString(inputFile);
       	 
       	 for(int ono = 1;ono < totNoOfOrders; ono++){ 
       	 if(nuberOfCells/2 < ono){
       		 System.out.println("not enough MSISDN and ICCID in excel");
       		 break;
       		 
       	 }
       	 
            String pattern = "(<so:SalesOrderLine>)(?s)(.*?)(</so:SalesOrderLine>)";
            Pattern r = Pattern.compile(pattern);
            Matcher salesOrderLineOccurances = r.matcher(testHtml);
            int count = 0;

   	         while(salesOrderLineOccurances.find()){
   	        	 count++;
   	        	 
   	        	 String pattern1 = "(<com:ID>)(?s)(.*?)(</com:ID>)";
   	        	 Pattern r1 = Pattern.compile(pattern1);
   	        	 Tex = new StringBuffer(salesOrderLineOccurances.group());
   	        	 Matcher idTagOccurances = r1.matcher(Tex);
   	    		 int count1 = 0;
   	        	 while(idTagOccurances.find()){
   	        		 count1++ ;

   	            	 int c = 1+(ono-1)*3;
   	        		 Integer c1 = new Integer(c);
   	        		 String c1String = c1.toString();
   	        		 Integer c2 = new Integer(c+1);
   	        		 String c2String = c2.toString();
   	        		 Integer c3 = new Integer(c+2);
   	        		 String c3String = c3.toString();
   	        		 
   	        		 String c1StringPattern = "<com:ID>"+ c1String + "</com:ID>";
   	        		 String c2StringPattern = "<com:ID>"+ c2String + "</com:ID>";
   	        		 String c3StringPattern = "<com:ID>"+ c3String + "</com:ID>";
   	        		 
   	        		 if(count == 1 && count1 == 1)
   	        			 Tex = Tex.replace(idTagOccurances.start(), idTagOccurances.end(),c1StringPattern);
   	        		 
   	        		 else if(count == 2 && count1 == 1)
   	        		 	Tex = Tex.replace(idTagOccurances.start(), idTagOccurances.end(),c2StringPattern);
   	        		 else if(count == 2 && count1 == 3)
   		        		 	Tex = Tex.replace(idTagOccurances.start(), idTagOccurances.end(),c1StringPattern);
   	        		 else if(count == 2 && count1 == 4){
   	        			String msisdnPattern = "<com:ID>"+ alMSISDN.get(ono) + "</com:ID>";
   	          		 	Tex = Tex.replace(idTagOccurances.start(), idTagOccurances.end(),msisdnPattern);    			 
   	        		 }
   	        		 else if(count == 3 && count1 == 1)
   		        		 	Tex = Tex.replace(idTagOccurances.start(), idTagOccurances.end(),c3StringPattern);
   	        		 else if(count == 3 && count1 == 2)
   		        		 	Tex = Tex.replace(idTagOccurances.start(), idTagOccurances.end(),c2StringPattern);
   	        		 else if(count == 3 && count1 == 3)
   		        		 	Tex = Tex.replace(idTagOccurances.start(), idTagOccurances.end(),c2StringPattern);
   	        		 else if(count == 3 && count1 == 5)
   		        		 	Tex = Tex.replace(idTagOccurances.start(), idTagOccurances.end(),c1StringPattern);
   	        		 else if(count == 3 && count1 == 6){
   	         			String iccidPattern = "<com:ID>"+ alICCID.get(ono) + "</com:ID>";
   	           		 	Tex = Tex.replace(idTagOccurances.start(), idTagOccurances.end(),iccidPattern);    			 
   	         		 }
   	        		 
   	        		 }
   	        	 	finalStringOfXML = finalStringOfXML.append(Tex);
   	        	 }

   	         }
          return finalStringOfXML;

         }finally {


         }

      }
    
//------------------------------------------    
    
    
    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(NewJFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(NewJFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(NewJFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(NewJFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new NewJFrame().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify                     
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTextArea jTextArea1;
    private javax.swing.JTextField jTextField1;
    // End of variables declaration                   
}
