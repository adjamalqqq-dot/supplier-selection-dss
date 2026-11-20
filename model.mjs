import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';
import fs from 'node:fs/promises';
const wb=Workbook.create(), sh=wb.worksheets.add('Supplier DSS');
sh.showGridLines=false;
sh.getRange('A1:I1').merge();sh.getRange('A1').values=[['SUPPLIER SELECTION  |  DECISION SUPPORT']];
sh.getRange('A2:I2').merge();sh.getRange('A2').values=[['Synthetic data • editable inputs in columns B–E and weights in B5:E5 • black text on white']];
sh.getRange('A4:F4').values=[['Weights','Cost','Quality','Delivery','Low risk','Weight check']];
sh.getRange('A5:E5').values=[['Editable',0.30,0.30,0.25,0.15]];
sh.getRange('F5').formulas=[['=IF(AND(COUNT(B5:E5)=4,MIN(B5:E5)>=0,MAX(B5:E5)<=1,ABS(SUM(B5:E5)-1)<0.000001),"OK","FIX WEIGHTS")']];
sh.getRange('A7:I7').values=[['Supplier','Cost RM','Quality /100','Delivery /100','Low risk /100','Cost score','Status','Total /100','Rank']];
sh.getRange('A8:E11').values=[['A',70,70,75,65],['B',90,90,90,80],['C',105,95,95,95],['D',120,85,80,90]];
for(let r=8;r<=11;r++){
 sh.getRange(`F${r}`).formulas=[[`=IF(AND(COUNT(B${r}:E${r})=4,MIN(B${r}:E${r})>=0,C${r}<=100,D${r}<=100,E${r}<=100),MAX(0,MIN(100,(130-B${r})/60*100)),"")`]];
 sh.getRange(`G${r}`).formulas=[[`=IF($F$5<>"OK","FIX WEIGHTS",IF(OR(A${r}="",COUNT(B${r}:E${r})<>4,MIN(B${r}:E${r})<0,C${r}>100,D${r}>100,E${r}>100),"FIX INPUT","OK"))`]];
 sh.getRange(`H${r}`).formulas=[[`=IF(COUNTIF($G$8:$G$11,"<>OK")=0,F${r}*$B$5+C${r}*$C$5+D${r}*$D$5+E${r}*$E$5,"")`]];
 sh.getRange(`I${r}`).formulas=[[`=IF(COUNTIF($G$8:$G$11,"<>OK")=0,1+COUNTIF($H$8:$H$11,">"&H${r}),"")`]];
}
sh.getRange('A14:I14').merge();sh.getRange('A14').values=[['INSTRUCTIONS']];
for(const [cell,text] of [['A15','1. Check and edit four weights in B5:E5; they must add to 100%.'],['A16','2. Edit supplier names and numeric inputs in B8:E11. Scores must be 0–100; cost is nonnegative.'],['A17','3. Lower price is better. Cost score = max(0, min(100, (130 − cost)/60 × 100)).'],['A18','4. Total = weighted sum of the four scores. Rank 1 is highest; equal totals share a rank.'],['A19','5. Invalid weights or supplier inputs suppress all totals and ranks. Fix every status before using results.'],['A20','6. Human review is essential. Scores, prices, weights and supplier names are fictional.'],['A21','Base weights: 30%, 30%, 25%, 15%. Cost-focused sensitivity: 60%, 15%, 15%, 10%.'],['A22','Prototype only: four fixed rows, no saved audit trail, no automated exclusion or final purchasing decision.']]) {sh.getRange(cell+':I'+cell.slice(1)).merge();sh.getRange(cell).values=[[text]];}
sh.getRange('A:I').format.columnWidth=17;sh.getRange('A:A').format.columnWidth=20;sh.getRange('G:G').format.columnWidth=20;sh.getRange('A1:I22').format.font.name='Arial';sh.getRange('A1:I22').format.font.size=11;sh.getRange('A1:I1').format.font.bold=true;sh.getRange('A1:I1').format.font.size=17;sh.getRange('A4:I4').format.font.bold=true;sh.getRange('A7:I7').format.font.bold=true;sh.getRange('A14:I14').format.font.bold=true;sh.getRange('A4:F5').format.borders={preset:'all',style:'thin',color:'#000000'};sh.getRange('A7:I11').format.borders={preset:'all',style:'thin',color:'#000000'};sh.getRange('A1:I22').format.rowHeight=23;sh.getRange('B5:E5').setNumberFormat('0%');sh.getRange('F8:F11').setNumberFormat('0.00');sh.getRange('H8:H11').setNumberFormat('0.00');
wb.recalculate();console.log((await wb.inspect({kind:'table',range:'Supplier DSS!A4:I11',include:'values,formulas',tableMaxRows:10,tableMaxCols:9})).ndjson);
const out=await SpreadsheetFile.exportXlsx(wb);await out.save('outputs/final_project/Supplier_Selection_DSS.xlsx');
try {const blob=await wb.render({sheetName:'Supplier DSS',range:'A1:I22',scale:1.5});await fs.writeFile('work/final_project/sheet.png',new Uint8Array(await blob.arrayBuffer()));} catch(e){console.log('render',String(e));}
