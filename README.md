# Cardiovascular-Risk-Calculator-in-Excel
This is a project i was assigned to in my jobplace. The Cardiovascular risk calculator computes the probability in percentage of dying or suffer a Cardiovascular negative event in certain groups of the population acoording to some health parameters.

These parameters are:

1.Gender. 
2.Age
3.systolic blood pressure
4.cholesterol(mg/dl).
5.smoker, non smoker.

Total of five (5) parameters.
Inside the Excel File there is a sheet named "TEST PAGE".There are excel cells identified with the headings corresponding to each parameter.Below this cell heading you must insert the value of each parameter and the function Calculate_Risk will compute the result.

The function RiskCV() can be called in any cell as long as the parameters are exactly identified and referenced in the formula editor. The function name is RiskCV(sex As String, age As Integer, smoke As String, PAS As Double, cholesterol As Double) As Double

This function "RiskCV" evaluates 400 conditions. Each condition is located inside the matrix. For each combination of the values of the parameters there is a result expressed in percentage.

This percentage represents the probability or the risk of the person to suffer a cardiovascular negative event or to face death by a cardiovascular accident.

The matrix of results which this excel project is based on is a documented study published by European Association for Cardiovascular Prevention & Rehabilitation (EACPR) and European Cardiology Society and also the European Atherosclerosis Society (EAS). This excel file is an implementation of the matrix of risk found in that document. This Study comes in the Guide ESC/EAS 2016. You can find this guide in http:\\www.revespcardiol.org

The matrix of results is included as a PNG file. The file is in Spanish.
The Excel file is a macro file(.xslm) made in the 2010 version of the software.
You can easily translate the few spanish terms into english.

In the excel file there is a sheet named "CONDITIONS" containing a table which enumerates every condition evaluated in the code.There is also a text file with the complete code in this repository.

You just have to download the files and test them in your PC. Scan files with antivirus for security before download or execution.

This project is intended to help health departments or any person involved in cardiovascular health issues.

