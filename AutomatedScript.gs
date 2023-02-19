// { 'Billed Unit': [ '0' ],
//   'Consumer Name': [ 'jksdnfv' ],
//   'Conn Load (in KW)': [ '0' ],
//   'L1 Hour': [ '0' ],
//   'L1 Days': [ '0' ],
//   'CA Number': [ '4343' ],
//   'L3 Factor': [ '0' ],
//   'L2 Hour': [ '0' ],
//   'Date of Inspection': [ '2/9/2023' ],
//   Supply: [ 'Domestic Supply' ],
//   'L4 Days': [ '0' ],
//   'L3 Hour': [ '0' ],
//   'Fixed Charge': [ '0' ],
//   'S Load (in KW)': [ '0' ],
//   'Email Address': [ 'ragib.hussain70@gmail.com' ],
//   Timestamp: [ '2/19/2023 22:43:42' ],
//   'L4 Hour': [ '0' ],
//   'Consumer Address': [ 'kjnsfjk' ],
//   'Meter Cost': [ '0' ],
//   'L2 Factor': [ '0' ],
//   'L3 Load': [ '0' ],
//   'L2 Load': [ '0' ],
//   'L1 Load': [ '0' ],
//   'L4 Factor': [ '0' ],
//   'L1 Factor': [ '0' ],
//   'L3 Days': [ '0' ],
//   'L4 Load': [ '0' ],
//   'L2 Days': [ '0' ] }

function onSubmit(e){
  
  const info = e.namedValues;

  console.log(info);
  const pdfFile = Create_PDF(info);
  
  sendEmail(e.namedValues['Email Address'][0],pdfFile,info);  
}

function getValues(){
  
  var slab0 = SpreadsheetApp.getActiveSheet().getRange('AG3').getValue();
  var slab1 = SpreadsheetApp.getActiveSheet().getRange('AG4').getValue();
  var slab2 = SpreadsheetApp.getActiveSheet().getRange('AG5').getValue();

  console.log(slab0, slab1, slab2);

}

function getL(load, factor, days, hour){

  return (load * factor * days * hour);

}

function sendEmail(email,pdfFile,info){
  
  const emailContent = "CA Number - " + info['CA Number'][0];
  GmailApp.sendEmail(email, "Provisional Assessment", emailContent, {
    attachments: [pdfFile], 
    name: "system-generated-mail"

  });
 
}
function Create_PDF(info) {
  
  const PDF_folder = DriveApp.getFolderById("1l02_MX6YbbmjTwci54gdvhrSH9AqKxK3");
  const TEMP_FOLDER = DriveApp.getFolderById("1-APGCaRPFByDtaTJr9ll_65RV8vF1zo5");
  const PDF_Template = DriveApp.getFileById("1ahYLAsm18BUZYKaJ3Y-jjkNXToirKFzcZ8XHCDs_LI0");
  
  const newTempFile = PDF_Template.makeCopy(TEMP_FOLDER);
  const OpenDoc = DocumentApp.openById(newTempFile.getId());
  const body = OpenDoc.getBody();
  
  console.log(body);
  
   body.replaceText("{supply}", info['Supply'][0]);
   body.replaceText("{caNum}", info['CA Number'][0]);
   body.replaceText("{dOI}", info['Date of Inspection'][0]);
   body.replaceText("{name}", info['Consumer Name'][0]);
   body.replaceText("{address}", info['Consumer Address'][0]);
   body.replaceText("{sLoad}", info['S Load (in KW)'][0]);
   body.replaceText("{cLoad}", info['Conn Load (in KW)'][0]);
   
   body.replaceText("{load1}", info['L1 Load'][0]);
   body.replaceText("{factor1}", info['L1 Factor'][0]);
   body.replaceText("{day1}", info['L1 Days'][0]);
   body.replaceText("{hour1}", info['L1 Hour'][0]);
   const L1 = getL(info['L1 Load'][0], info['L1 Factor'][0], info['L1 Days'][0], info['L1 Hour'][0])
   body.replaceText("{value1}", L1);

   body.replaceText("{load2}", info['L2 Load'][0]);
   body.replaceText("{factor2}", info['L2 Factor'][0]);
   body.replaceText("{day2}", info['L2 Days'][0]);
   body.replaceText("{hour2}", info['L2 Hour'][0]);
   const L2 = getL(info['L2 Load'][0], info['L2 Factor'][0], info['L2 Days'][0], info['L2 Hour'][0]);
   body.replaceText("{value2}", L2);

   body.replaceText("{load3}", info['L3 Load'][0]);
   body.replaceText("{factor3}", info['L3 Factor'][0]);
   body.replaceText("{day3}", info['L3 Days'][0]);
   body.replaceText("{hour3}", info['L3 Hour'][0]);
   const L3 = getL(info['L3 Load'][0], info['L3 Factor'][0], info['L3 Days'][0], info['L3 Hour'][0]);
   body.replaceText("{value3}", L3);

   body.replaceText("{load4}", info['L4 Load'][0]);
   body.replaceText("{factor4}", info['L4 Factor'][0]);
   body.replaceText("{day4}", info['L4 Days'][0]);
   body.replaceText("{hour4}", info['L4 Hour'][0]);
   const L4 = getL(info['L4 Load'][0], info['L4 Factor'][0], info['L4 Days'][0], info['L4 Hour'][0]);
   body.replaceText("{value4}", L4);

   const TAU = L1 + L2 + L3 + L4;
   body.replaceText("{totalValue}", TAU);

   body.replaceText("{billedUnit}", info['Billed Unit'][0]);

  OpenDoc.saveAndClose();
  

  const BLOBPDF = newTempFile.getAs(MimeType.PDF);
  const pdfFile =  PDF_folder.createFile(BLOBPDF).setName(info['Consumer Name'][0] + " " + info['CA Number'][0]);
  console.log("The file has been created ");
  
  return pdfFile;

}
