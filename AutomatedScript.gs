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

function Tester(){
  getEC("Domestic", 15153.72, 1272, 12);
  //getFixedCharge("Domestic", 12, 1, 2.472)
}

function getSlab(supply){
  
  var slab0, slab1, slab2;

  if(supply == "Domestic"){
    slab0 = SpreadsheetApp.getActiveSheet().getRange('AD3').getValue();
    slab1 = SpreadsheetApp.getActiveSheet().getRange('AD4').getValue();
    slab2 = SpreadsheetApp.getActiveSheet().getRange('AD5').getValue();
  }
  else{
    slab0 = SpreadsheetApp.getActiveSheet().getRange('AD8').getValue();
    slab1 = SpreadsheetApp.getActiveSheet().getRange('AD9').getValue();
    slab2 = SpreadsheetApp.getActiveSheet().getRange('AD10').getValue();
  }

  console.log("Printing Slab - ", slab0, slab1, slab2);

  const slab = [slab0, slab1, slab2];
  return slab;

}

function getFactor(supply, status){

  var factor = 0;

  if(status == "Unmetered"){
    factor = 1;
    return factor;
  }

  if(supply == "Domestic"){
    factor = 0.3;
  }
  else{
    factor = 0.5;
  }

  return factor;

}

function getSlabbedEC(unit, months){
  
  var val0 = 0;
  var val1 = 0;
  var val2 = 0;

  if(unit <= 100 * months){
    val0 = unit;  
  }
  else{
    val0 = 100 * months;
    unit = unit - (100 * months);
    if(unit <= 100 * months){
      val1 = unit;
    }
    else{
      val1 = 100 * months;
      unit = unit - (100 * months);
      val2 = unit;
    }
  }
  console.log(val0, val1, val2);

  const slabWiseCharges = [val0, val1, val2];
  
  return slabWiseCharges;
}

function getL(load, factor, days, hour){

  return (load * factor * days * hour);

}

function getEC(supply, TAU, BU, months){

  console.log(supply, TAU, BU, months);

  const slab = getSlab(supply);

  const totalEC = getSlabbedEC(TAU, months);
  const paidEC = getSlabbedEC(BU, months);

  const diff0 = totalEC[0] - paidEC[0];
  const diff1 = totalEC[1] - paidEC[1];
  const diff2 = totalEC[2] - paidEC[2];

  const amount0 = diff0 * slab[0];
  const amount1 = diff1 * slab[1];
  const amount2 = diff2 * slab[2];

  const finalAmount = amount0 + amount1 + amount2;

  const EC = [totalEC[0], totalEC[1], totalEC[2], paidEC[0], paidEC[1], paidEC[2], slab[0], slab[1], slab[2], diff0, diff1, diff2, amount0, amount1, amount2, finalAmount];

  console.log("Printing EC - ", EC);
  return EC;

}

function getFixedChargeSlab(supply){

  var slab = 0;

  if(supply == "Domestic"){
    slab = SpreadsheetApp.getActiveSheet().getRange('AD13').getValue();
  }
  else{
    slab = SpreadsheetApp.getActiveSheet().getRange('AD16').getValue();
  }

  console.log(slab);

  return slab;
}

function getDiffLoad(sLoad, cLoad){

  const diffLoad = Math.ceil(cLoad - sLoad);
  console.log(diffLoad);

  return diffLoad;
}

function getFixedCharge(supply, months, sLoad, cLoad){

  console.log(supply, months, sLoad, cLoad);

  const fixedChargeSlab = getFixedChargeSlab(supply);
  const diffLoad = getDiffLoad(sLoad, cLoad);
  const fixedCharge = fixedChargeSlab * months * diffLoad;

  console.log(fixedCharge);
  return fixedCharge;

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

  const months = (info['L1 Days'][0] / 365) * 12;
  const lFactor = getFactor(info['Category'][0], info['Consumer Status'][0]);

  
  console.log(body);
  
   body.replaceText("{supply}", info['Category'][0]);
   body.replaceText("{conStatus}", info['Consumer Status'][0]);
   body.replaceText("{caNum}", info['CA Number'][0]);
   body.replaceText("{dOI}", info['Date of Inspection'][0]);
   body.replaceText("{name}", info['Consumer Name'][0]);
   body.replaceText("{address}", info['Consumer Address'][0]);
   body.replaceText("{sLoad}", info['Sanctioned Load (in KW)'][0]);
   body.replaceText("{cLoad}", info['Connected Load (in KW)'][0]);
   
   body.replaceText("{load1}", info['L1 Load'][0]);
   body.replaceText("{factor1}", lFactor);
   body.replaceText("{day1}", info['L1 Days'][0]);
   body.replaceText("{hour1}", info['L1 Hour'][0]);
   const L1 = getL(info['L1 Load'][0], lFactor, info['L1 Days'][0], info['L1 Hour'][0])
   body.replaceText("{value1}", L1);

   body.replaceText("{load2}", info['L2 Load'][0]);
   body.replaceText("{factor2}", lFactor);
   body.replaceText("{day2}", info['L2 Days'][0]);
   body.replaceText("{hour2}", info['L2 Hour'][0]);
   const L2 = getL(info['L2 Load'][0], lFactor, info['L2 Days'][0], info['L2 Hour'][0]);
   body.replaceText("{value2}", L2);

   body.replaceText("{load3}", info['L3 Load'][0]);
   body.replaceText("{factor3}", lFactor);
   body.replaceText("{day3}", info['L3 Days'][0]);
   body.replaceText("{hour3}", info['L3 Hour'][0]);
   const L3 = getL(info['L3 Load'][0], lFactor, info['L3 Days'][0], info['L3 Hour'][0]);
   body.replaceText("{value3}", L3);

   body.replaceText("{load4}", info['L4 Load'][0]);
   body.replaceText("{factor4}", lFactor);
   body.replaceText("{day4}", info['L4 Days'][0]);
   body.replaceText("{hour4}", info['L4 Hour'][0]);
   const L4 = getL(info['L4 Load'][0], lFactor, info['L4 Days'][0], info['L4 Hour'][0]);
   body.replaceText("{value4}", L4);

   const TAU = L1 + L2 + L3 + L4;
   body.replaceText("{totalValue}", TAU);

   body.replaceText("{billedUnit}", info['Billed Unit'][0]);

   body.replaceText("{chargeableUnit}", (TAU - info['Billed Unit'][0]));

   // Computing energy charge based on slab, for domestic supply (DS) and non domestic supply (nds) different rates will be applied.
   const energyCharge = getEC(info['Category'][0], TAU, info['Billed Unit'][0], months);
   body.replaceText("{totalEC0}", energyCharge[0]);
   body.replaceText("{totalEC1}", energyCharge[1]);
   body.replaceText("{totalEC2}", energyCharge[2]);
   body.replaceText("{paidEC0}", energyCharge[3]);
   body.replaceText("{paidEC1}", energyCharge[4]);
   body.replaceText("{paidEC2}", energyCharge[5]);
   body.replaceText("{slab0}", energyCharge[6]);
   body.replaceText("{slab1}", energyCharge[7]);
   body.replaceText("{slab2}", energyCharge[8]);
   body.replaceText("{diff0}", energyCharge[9]);
   body.replaceText("{diff1}", energyCharge[10]);
   body.replaceText("{diff2}", energyCharge[11]);
   body.replaceText("{amount0}", energyCharge[12]);
   body.replaceText("{amount1}", energyCharge[13]);
   body.replaceText("{amount2}", energyCharge[14]);
   body.replaceText("{energyCharge}", energyCharge[15]);

   // Applying 6% additional charge to energy charge.
   const ed = 0.06 * energyCharge[15];
   body.replaceText("{ed}", ed);

   const total = energyCharge[15] + ed;
   body.replaceText("{total}", total);

   const fixedCharge = getFixedCharge(info['Category'][0], months, info['Sanctioned Load (in KW)'][0], info['Connected Load (in KW)'][0]);
   body.replaceText("{rate}", getFixedChargeSlab(info['Category'][0]));
   body.replaceText("{month}", months);
   body.replaceText("{diff}", getDiffLoad(info['Sanctioned Load (in KW)'][0], info['Connected Load (in KW)'][0]));
   body.replaceText("{fcValue}", fixedCharge);

   const grossTotal = total + fixedCharge;
   body.replaceText("{grossTotal}", grossTotal);

   const twiceRate = 2 * grossTotal;
   body.replaceText("{twiceRate}", twiceRate);

   body.replaceText("{meterCost}", info['Meter Cost'][0]);

   const meterGST = 0.18 * info['Meter Cost'][0];
   body.replaceText("{gst}", meterGST);

   const totalMeterCost = (1 * info['Meter Cost'][0]) + meterGST;
   body.replaceText("{totalMeterCost}", totalMeterCost);
   
   const finalPenalty = twiceRate + totalMeterCost;
   body.replaceText("{finalPenalty}", finalPenalty);

  OpenDoc.saveAndClose();
  

  const BLOBPDF = newTempFile.getAs(MimeType.PDF);
  const pdfFile =  PDF_folder.createFile(BLOBPDF).setName(info['Consumer Name'][0] + " " + info['CA Number'][0]);
  console.log("The file has been created ");
  
  return pdfFile;

}
