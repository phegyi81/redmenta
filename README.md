# redmenta
How to dump Redmenta exams at upload them to Edubase

Dumping Redmenta to Edubase exercieses is built up from the following phases.

1. Phase: save the Redmenta exams. Open the editor page, DO NOT edit any of the exercieses and save the HTML. The complete webpage is needed, otherwise the exercieses will be missing somehow.

2. Phase: `node html-parser.js`
   - This will get out the exercieses from ALL HTML file in the directory and creates an Excel file (NOT CSV!) to each of them.

3. Phase: fill in the solutions in the second column of the Excel file
   -numbers: the number of options being valid - multiple answers may valid!
   -r: short answer
   -k: long answer
   -a series of IH: true or false questions. Be careful: the order is not the same as appearing in the exam!
   -m: matching question. Anwsers are in order, that is the first half is the first column, the second half is the second column and the options in the same positions of the two columns are matching.

4. Phase: convert the Excel files to be in Edubase format
   -Do not forget to randomize option order!!!
   
Not handled!
- If more than one point can be earned in an exerciese.
