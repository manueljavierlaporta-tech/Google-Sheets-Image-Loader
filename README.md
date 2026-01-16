# 🖼️ Google Sheets | Image Loader

This project contains a **Google Apps Script** exntension for **Google Sheets**, designed to load external images (in this particular case, is specific for a type of Shutterstock URLs) into a spreadsheet while preserving metadata and binding the image to a specific row, using native Google Sheets formulas and, of course, Google App Scripts (so, JavaScript).  

The solution uses a custom HTML modal to collect data and a backend App Script function to insert rows programmatically.

## ✍🏻 Workflow

<div>
  <p> This scriptr was created to support the registration of Shutterstock photo purchases for an <b>advertising agency</b>. The creative assets (especially the image code and preview) need to be easily trackable (hello Ctrl+F), and the agency needs to be able to preview the image to decide which one to use.<br>
The process works as follows: </p>
  <ol>
    <li> A user opens a custom menu inside Google Sheets. </li>
    <li> A modal dialog (HTML + CSS + JavaScript) is displayed. </li>
    <li> The user pastes one or more image URLs (e.g. Shutterstock links) and related descriptive data. </li>
    <li> The script extracts:
      <ul>
        <li> Image ID </li>
        <li> Insertion date </li>
        <li> Description / notes </li>
        <li> Internal reference code (if present) </li>
        <li> Image formula bound to a specific cell </li>
      </ul>
    </li>
    <li> A new row is appended for each image using Apps Script. </li>
  </ol>
</div>

## 🎯 Script Objective

The backend App Script must:  
<div>
  <ol>
    <li> Receive structured data from an HTML dialog. </li>
    <li> Process multiple URLs in a single submission. </li>
    <li> Extract identifiers from each URL. </li>
    <li> Generate a valid <code>=IMAGE()</code> formula. </li>
    <li> Append the data as a new row in the active sheet. </li>
  </ol>
</div>

<p> The HTML dialog sends data to AppScript using <code>google.script.run</code>.<br>
Each image URL is processed and converted into a row:</p>
```
// Finding Shutter image code from URL
for(let k=0; k < urlValuesArray.length; k++){
  let shutterstockCode = urlValuesArray[k].split("/")[5];
  let today = `${ new Date().getDate() }/${ parseInt( new Date().getMonth() ) + 1 }/${ new Date().getFullYear() }`;
  let about = aboutValuesArray[k] || "";
  let internalCode = about.includes("BM") ? about.split(" ")[0] : "";
  let fullUrl = `=image("${urlValuesArray[k]}")`;

  let allValuesArray = [
    shutterstockCode,
    today,
    about,
    internalCode,
    fullUrl
  ];
    
  google.script.run.newRow(allValuesArray);
};
```
