<h2>Introduction</h2>
 
 <p>This .NET Framework desktop application can read a WSDL XML for EJB SOAP APIs, extract information from them, and summarize and organize the data on a word document. This application cannot describe web services generated from .NET, which I may tackle later in the future, and this application cannot describe all Java-based web services, since I need more samples to adjust the source code accordingly.<br/>

This application was developed in 2015, as a result of developing a similar application for a previous company to save time on writing a well written integration guide for EJB web services. I hope that, by making it open source, the application would work a lot better and more beneficial towards others, even though more and more SOAP APIs are becoming less popular.
   
 </p>

<h2>Implementation</h2>
<p>Other than the Main class, there are ten classes in the project:</p>

<ul>
  <li>ApplicationManager</li>
  <li>MainUI</li>
  <li>WebMethodInfo</li>
  <li>WordWriter</li>
  <li>WsdlOperation</li>
  <li>WsdlPortType</li>
  <li>XSComplexType</li>
  <li>XSElement</li>
  <li>XSSchema</li>
  <li>XSSimpleType</li>
</ul>

<p>Now before we get into the classes that controls the work flow of the application, I want to explain about the classes that is used to hold the element data within the WSDL. Let’s call them the “WSDL classes”.</p>

<h3>The WSDL Classes</h3>

<p>If you are unfamiliar on how WSDL is read and interpreted, please check out the tutorial <a href="https://www.predic8.com/wsdl-reading.htm">here.</a> It’s pretty useful.<br/>

Other than WebMethodInfo, the classes that are considered to be the WSDL classes are WsdlOperation, WsdlPortType, XSComplexType, XSElement, XSSchema and XSSimpleType. The names of the classes are based on the names of the XML element types in the WSDLs. For instance, XSElement represents , XSComplexType represents , and WsdlOperation represents , and so on. When reading the XML, the elements info are stored in these classes in order to interpret them when writing documents and filling in the WebMethodInfo object.<br/>

You might ask on why I didn’t write the document and fill the WebMethodInfo object while getting the info from the WSDL XML. Well, it is incredibly difficult and time consuming to do so, and would potentially increase the time complexity of the program, making it potentially slower.</p>

<h3>MainUI Class</h3>

<p>As the name of the class may suggest, this class is used to create and control the graphical user interface that the user will interact with and call the application manager. There are six windows application controls that the user will be concerned with:</p>

<ul>
  <li>urlBox (Textbox): This text box is used to insert the URL of the WSDL XML to generate the description document from.</li>
<li>authorNameBox (Textbox): This text box is used to insert the name of the author of the document (the user of WSDL Describer) in the generated document.</li>
<li>statusBox (RichTextBox): This box will display the written status of generating the WSDL Description document. For example, if it is extracting info from the WSDL XML, it will display “Extracting WSDL information from XML”.</li>
<li>generatorButton (Button): This button will call the save file dialog. If the user chooses the directory of where the newly created Word document will be stored and clicks the “Ok” button, the program will begin the process.</li>
<li>saveFileDialog (SaveFileDialog): The dialog box will allow the user to choose the directory of where the newly created Word document will be stored.</li>
<li>progressBar (ProgressBar): This bar will show how far along the process of generating the document is coming along. The longer the green area in the progress bar is, the closer it is to completing generating the document.</li>
</ul>

![Describing_WSDLDescriber_UI](https://github.com/HussamAlnasser/WSDLDescriber/assets/12647832/3c647d9f-dd5b-42fe-87ff-e7e13650553e)

<p>There is also one more control in the MainUI class that the user does not interact with, but it affects the other controls.</p>

<ul><li>timer (Timer): This control is used to update the status of generating documents. This is done by creating a new thread after choosing “OK” in the saveDialogBox to generate the document, and use the main thread to call the timer method in order to update the progressBar and statusBox.</li></ul>

![WSDLDescriberMainUICode](https://github.com/HussamAlnasser/WSDLDescriber/assets/12647832/e6a1dd15-9e55-4a80-88b1-abf2c8e75125)

<h3>ApplicationManager Class</h3>

<p>This is the class that calls all of the methods crucial to the process of generating the WSDL Description document. There are three methods in this class, and one of which calls the other two methods:</p>

<ul>

  <li>ExtractWSDLInfo: This method extracts information from the WSDL XML, and stores it in the WSDL classes mentioned above.</li>
  <li>FillInWebMethodInfo: This method takes the extracted information stored in the WSDL objects and use them to interpret the information and store the interpretation in the WebMethodInfo objects for each web method in the web service application</li>
  <li>DescribeWSDLInDocument: This is where all of the data flow to generate the document happens. It will call ExtractWSDLInfo, FillInWebMethodInfo, and the WriteDocument in the WordWriter class.</li>
</ul>

<h3>WordWriter Class</h3>

<p>This is the class that will write the document. The main method of WordWriter that will be called in the ApplicationManager class is WriteDocument. However, many of the logic of the method is divided into manyh other methods in order to reduce the time complexity and organizing the section of the document.

The most important thing to notice in the class is the repeating lines of code crucial to writing the paragrtaphs. They are:</p>

<code>newParagraph = wordDoc.Content.Paragraphs.Add();</code>

<p>which creates a new paragraph, and:</p>

<code>newParagraph.Range.InsertParagraphAfter();</code>

which ends the paragraph instantly.

<p>The class is divided into nine methods:</p>

<ul>
  <li>CreateCoverPage: This creates the cover page of the word document</li>
  <li>CreateIntroduction: This creates the introduction section of the document</li>
  <li>CreateListTemplate: This creates the list format in the document. So, this is where you can customize the way the list is organzined</li>
  <li>CreateXSD: This method creates the XML Schema Definition mini section in the Request and Response Message sections. This method will also call the CreateXSDString method to get the XSD string that defines the Request and Response Messages in the WSDL XML.</li>
  <li>CreateXSDString: This method will get the XSD string that defines the Request and Response Messages in the WSDL XML.</li>
  <li>CreateFieldDescription: This will create the tables that list the variable that are used as inputs and outputs. This method will also call the AddDataToFieldDecription to fill in the data and add rows in the tables</li>
  <li>AddDataToFieldDecription: This method will fill in the data in the Field Description tables. It will call itself recursively in order to add extra rows and data when necessary.</li>
  <li>CreateWSDLAppendix: This will write all of the raw WSDL XML string at the end of the document.</li>
  <li>WriteDocument: This is the main method that is used to write the document and call all of the mentioned methods above</li>
</ul>
