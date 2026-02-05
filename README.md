# IMT-Atlantique-NEPIA-M1S1-Scientific_Project
<p>Txt-to-Spreadsheet Scripts</p>
<p>By: Javier. https://sites.google.com/view/b-eng-jarl/home</p>
<br></br>
<p>Serpent by VTT is a big-data processing tool used to simulate nuclear engineering principles.</p>
<p>A single instance of Serpent usage resulted in up to 110 data points reported in an unstructured, human-readable only, terminal-based manner.</p>
<p>The following scripts were created to assist the filtering and orderly listing of such data points in the interest of extending Serpent reported data points into first-principle-level Machine Learning models.</p>
<p></p>
<style>
.myDiv {
  border: 5px outset red;
  background-color: lightblue;
  text-align: center;
}
</style>
<div class="myDiv">
  <h2>keffReader.py :</h2>
  <p>Description: This script takes terminal outputs in *.txt format and scans for Serpent output data points. These are then ordered and listed into a new spreadsheet workbook.</p>
  <p>Usage: User is to copy-paste Serpent terminal reports into a *.txt file. Then, the user is to input the *.txt file when prompted using local OS-native GUIs.</p>
</div>
<div class="myDiv">
  <h2>mergeODS.py :</h2>
  <p>Description: This script merges keffReader.py outputs into a single workbook *.ods.</p>
  <p>Usage: User is to select the keffReader.py out files when prompted using local OS-native GUIs.</p>
</div>

