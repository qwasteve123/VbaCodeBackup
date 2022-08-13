# VbaLinkBudBackup

This file is for design and management of link budget of Indoor distributed control system (DCS)  network. The aim for this project is to automate the tedious calculations in designing routing of feeder cable and choice of coupler in the whole rf network.

 ## Content

1. [Introduction](#md-header2-introduction)
2. [User Manual](#md-header2-user-manual)
    1. [Set up visio file](#md-header3-set-up)
    2. [(Optional) Set up shortcut key for macro](#md-header3-fast-key)
    3. [How To Use It](#md-header3-how-to-draw)
    4. [Edit Shape in Shape Data](#md-header3-edit-shape-data)
    5. [Fast Naming of Shapes](#md-header3-fast-naming)
    6. [Change Label Format](#md-header3-label-format)
    7. [Show or Export Link Budget](#md-header3-link-bud)


<h2 id="md-header2-introduction">Introduction </h2>
The main problem in indoor DCS system is the complexity in calculation to reach a balance of power output in every antenna. The system has a source to distrubute to every descending separation point (couplers) and eventually reach the output (antenna). The complexity of DCS system design comes from the limitation in construction site, mainly the pathway restricted and the antenna location pre-requested in tender by clients. To ensure the calculations reach the demanded power output (RSRP), link budget calculation is required. Without automation, 
we need to care about the choice of couplers, length and routing of feeder cables. Changing a coupler will alter the descendents' output which consisted in the same routing, which makes link budgets design difficult and complex.<br><br>
  
This project focus on calculation of RSRP output from antenna by subtracting the coupling loss and feeder loss in the whole pathway from source to ouput antenna. Visio is used for showcase of link budget and DCS network design. 

<h2 id="md-header2-user-manual"> User Manual </h2>

<h3 id="md-header3-set-up"> Set up visio file </h3>
If you already have a visio file with this code, just delete all contents in the document. The document is ready for use. <br><br>

1. For setting up visio document, you need to open Visio first. Press Development tab in toolbar and you can see Visual Basic Editor (VBE). Press it to enter VBE.
<br>___image required___
2. Download all files from this project, drag all VBA files (___.bas, .frm, .cls file only___) to VB Editor as follow:
<br>___image required___
3. Close VB Editor. Back to Visio file.
4. Choose ___File >>> Shapes >>> Open Stensils...,___ a file dialog will pop up. Choose ___.vss file___ in this project.
<br>___image required___
5. You can see the stensil with shapes shown up. It is recommended to put all shapes to document stensil. The document is set up.<br>
<br>___check if need to make dynamic connector___

<h3 id="md-header3-fast-key"> (Optional) Set up shortcut for macro </h3>

1. Press run macro, from ___Macro in:___ dropdown list, select **A_main**
<br>___image required___
2. Select __ShowLinkBudget__ from list, Press __Option...__ to assign shortcut key. Press ___Shift + D___ and press __Ok__.
<br>___image required___
3. Select __ExportLinkBudget__ from list, Press __Option...__ to assign shortcut key. Press ___Shift + F___ and press __Ok__.
4. Back to __Macro in:__ dropdown list, select __DC_AutoNumbering__
<br>___image required___
5. Select __ActivateForm__ from list, Press __Option...__ to assign shortcut key. Press __Shift + W__ and press __Ok__. 

<h3 id="md-header3-how-to-draw"> How To Draw Schametic Diagram</h3>

1. The component shapes follow the conventional shape of couplers, splitter and antennae in schematic diagram. DCS system start from ___Start Block___ to ___Coupler/
2 way splitter/ 3 way splitter___ to ___Antenna___ and linked by ___Feeder cable___ in connection points. Connection method is based on schematic drawing.
2. To place shapes, press and drag shape from ___Document Stensil___ to your page.
<br>___image required___
3. To connect shapes, press ___Ctrl + 3___ to call dyamic connector and connect each shape.
 <br>___image required___
 
 <h3 id="md-header3-edit-shape-data"> Edit Shapes in Shape Data</h3>
 
1. To change the choice of couplers, right click on shape and click ___shape data___. Change coupler in ___coupling loss___ dropdown list.
<br>___image required___
2. To change the choice of feeder cable, right click on shape and click ___shape data___. Change feeder cable in ___feeder type___ dropdown list.
<br>___image required___
3. To change the antenna gain, right click on master shape in document stensil, click ___edit master___ . After the master shape is opened,right click on master shape and click ___shape data___. Change your antenna gain.
<br>___image required___
4. To change shape floor and item no. shown on shape text, edit __shape data___ window. Be careful that item no. can only accept ___integer___ while floor can be any text (string datatype). For more convinient method, read section [Fast Naming of Shapes](#md-header3-fast-naming)
<br>___image required___

 <h3 id="md-header3-fast-naming"> Fast Naming of Shapes</h3>
 
1. To change floor of multiple shapes, select multiple shapes, right click on shape and click ___shape data___ to change floor. Do not include shape that without floor input in your selection otherwise floor properties will not be shown in shape data window.
2. There are two methods to update item no. of shape:
     1. To name shapes item no. in consecutive order ( 1,2,3,...), Press ___Ctrl + Shift + W___ to active subroutine, ActiveForm. (you can also run it in ___run       macro___). Go to tab ___Consecutive___, Press ___Start___ button to start naming. Click the shape you want to name, the number will increase in increment of 1.  
      To change the number added to shapes, enter new number in the textbox on left and click ___Change Number___ and continue clicking shapes.
      <br>___image required___
     2. Sometimes you need to update all shapes in increment of 1. In this situation, Press ___Ctrl + Shift + W___ to active subroutine, Go to tab ___Add number___,
      Enter the increment you want and press ___Selections Add___.
      <br>___image required___

<h3 id="md-header3-label-format"> Change Label Format</h3>
This project consists of the two label format, Normal label format and Lift label format. To switch between two label format, select multiple shapes, choose to ___Lift label format___ and click ___Change Label Format___.

<br>___image required___

 <h3 id="md-header3-link-bud"> Show or Export Link Budget</h3>

 
 
