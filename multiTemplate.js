function getMultiTemplate(templateID){
  var templateDeck = SlidesApp.openById(templateID);
  var templateSlides = templateDeck.getSlides();

  const categories = [
    "Arts & culture",
    "Music",
    "Science & Technology",
    "Sports",
    "Others"
  ];

  let templateArr = categories.reduce(
    (o, c, i) => Object.assign(o, { [c]: templateSlides[i]}), // o: object, c: category, i: index
    {}
  );


  // let templateArr = [];
  // templateArr.push({ ["Arts & culture"]: templateSlides[0]});
  // templateArr.push({ ["Music"]: templateSlides[1]});
  // templateArr.push({ ["Science & Technology"]: templateSlides[2]});
  // templateArr.push({ ["Sports"]: templateSlides[3]});
  // templateArr.push({ ["Others"]: templateSlides[4]});

  // templateArr.push({"category": "Arts & culture", "template": templateSlides[0]});
  // templateArr.push({"category": "Music", "template": templateSlides[1]});
  // templateArr.push({"category": "Science & Technology", "template": templateSlides[2]});
  // templateArr.push({"category": "Sports", "template": templateSlides[3]});
  // templateArr.push({"category": "Others", "template": templateSlides[4]});

  // console.log(templateArr);

  return [categories, templateArr];
}

function main2() {
  // ID of the slidess
  var masterDeckID = "1hsRpBKQ1wGbwsliqLCEDYhw-8-2oGzUhK8HW22r_gM0"; // Circlepedia Generated Slides
  var templateID = "1HLQFQY296I3MUkinDFjuf5MmCKZOf8GQse1nxq_jT6E"; // Circlepedia Slideshow Templates

  // Open the presentation
  var masterDeck = SlidesApp.openById(masterDeckID);
  var slides = masterDeck.getSlides();

  // Get the width and height of the presentation
  // var pageWidth = masterDeck.getPageWidth();
  // var pageHeight = masterDeck.getPageHeight();

  // Remove everything starting from (including) slide with Index 1
  const nIndex = 1;
  slides.slice(nIndex).forEach(s => s.remove());

  // Add the template slide to the masterDeck
  var [categories, templateArr] = getMultiTemplate(templateID);
  // templateArr.forEach((t) => masterDeck.appendSlide(t.template));
  categories.forEach((category) => {
    templateArr[category] = masterDeck.appendSlide(templateArr[category]);
  });
  // masterDeck.getMasters()[0].remove(); // remove original theme or sth (to copy the theme of the template)

  // slides = masterDeck.getSlides();

  // template = slides[1];

  // Load data from spreadsheet. Active spreadsheet means the one open right now.
  // getDataRange() is functionally equivalent to creating a Range bounded by A1 and (Sheet.getLastColumn(), Sheet.getLastRow()).
  var dataRange = SpreadsheetApp.getActive().getDataRange();
  var sheetContents = dataRange.getValues(); // returns 2D array

  // Remove the header (1st row)
  sheetContents.shift();

  // Reverse the order of the rows (because newer slides will be inserted at the top)
  sheetContents.reverse();

  var counter = 0;
  var arr = [];

  sheetContents.forEach(function (row) {

    console.log('Slide number: ', counter);
    counter = counter + 1;

    var fields = [
      ["{{Name (Jap)}}", row[2]],
      ["{{Name (Eng)}}", row[3]],
      ["{{Category}}", row[4]],
      ["{{Activities}}", row[5]],
      ["{{Date}}", row[6]],
      ["{{Place}}", row[7]],
      ["{{Fee}}", row[8]],
      ["{{Eligibility}}", row[9]],
      ["{{Japanese}}", row[10]],
      ["{{Contact}}", row[11]]
      ];

    // Insert a new slide by duplicating the template slide.
    let slide = templateArr[row[4]].duplicate();

    arr.push({"name": row[3], "slideRef": slide}); // insert English circle name & current slide

    
    console.log("Current slide: ", row[2]);
    fields.forEach(f => slide.replaceAllText(...f)); // spread operator 

    let website = row[14]; // e.g. "https://juggling-donuts.org/"
    let twitterHandle = row[15]; // e.g. "@soajo_KUMC"
    let igHandle = row[16]; // e.g. "@soajo_kumc"
    let fbHandle = row[17]; // e.g. "iGEM Kyoto"

    const shapes = slide.getShapes();
    if (shapes.length > 0) {
      shapes.forEach(shape => {
        processShape({
          "shape": shape,
          "website": website,
          "twitterHandle": twitterHandle,
          "igHandle": igHandle,
          "fbHandle": fbHandle
        });
      });
    }

    const processGroups = g => {
      g.getChildren().forEach(c => {
        const type = c.getPageElementType();
        if (type == SlidesApp.PageElementType.SHAPE) {
          processShape({
            "shape": c.asShape(),
            "website": website,
            "twitterHandle": twitterHandle,
            "igHandle": igHandle,
            "fbHandle": fbHandle
          });
        } else if (type == SlidesApp.PageElementType.GROUP) {
          processGroups(c.asGroup());
        }
      });
    }
    slide.getGroups().forEach(processGroups);

    function processShape(arg) {
      let website = arg.website;
      let twitterHandle = arg.twitterHandle;
      let igHandle = arg.igHandle;
      let fbHandle = arg.fbHandle;

      let text = arg.shape.getText();
      switch (true) {
        case text.asString().includes("{{Website}}"): // Website
          if (website) {
            let n = text.replaceAllText("{{Website}}", website);
            if (n > 0) {
              text.find(website).forEach((v) => {v.getTextStyle().setLinkUrl(website)});
            }
          } else {
            text.replaceAllText("{{Website}}", "/");
          }

        case text.asString().includes("{{Twitter}}"): // Twitter
          if (twitterHandle){
            let twitterUsername = twitterHandle.split('@')[1]; // remove the @
            let twitterUrl = "https://twitter.com/" + twitterUsername;
            let n = text.replaceAllText("{{Twitter}}", twitterHandle);
            if (n > 0) {
              text.find(twitterUsername).forEach((v) => {v.getTextStyle().setLinkUrl(twitterUrl)});
            }
          } else {
            text.replaceAllText("{{Twitter}}", "/");
          }
        case text.asString().includes("{{Instagram}}"): // Instagram
          if (igHandle){
            let igUsername = igHandle.split('@')[1]; // remove the @
            let igUrl = "https://www.instagram.com/" + igUsername;
            let n = text.replaceAllText("{{Instagram}}", igHandle);
            if (n > 0) {
              text.find(igUsername).forEach((v) => {v.getTextStyle().setLinkUrl(igUrl)});
            }
          } else {
            text.replaceAllText("{{Instagram}}", "/");
          }
        case text.asString().includes("{{Facebook}}"): // Facebook
          if (fbHandle){
            let fbUsername = fbHandle.replace(/\s+/g, ''); // remove all spaces
            let fbUrl = "https://www.facebook.com/" + fbUsername;
            let n = text.replaceAllText("{{Facebook}}", fbHandle);
            if (n > 0) {
              text.find(fbHandle).forEach((v) => {v.getTextStyle().setLinkUrl(fbUrl)});
            }
          } else {
            text.replaceAllText("{{Facebook}}", "/");
          }
        default:
          // console.log("Do nothing");
      }
    }

    //Logo image
    if (row[12] != ""){
      // console.log('type of row[12]', typeof row[12]);
      // console.log('row[12]', row[12]);
      let regExp = new RegExp("[^=]*$");
      let logoId = regExp.exec(row[12])[0];

      var logo = UrlFetchApp.fetch(`https://drive.google.com/thumbnail?sz=w1000&id=${logoId}`, { headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() } }).getBlob(); // The endpoint is from https://stackoverflow.com/a/31504086
      // console.log(typeof logoId);
      // console.log('flyerid: ', logoId);

      // let logo = DriveApp.getFileById(logoId);

      // slide.insertImage(image);
      slide.getShapes().forEach(shape => {
        if (shape.getText().asString().trim() == "{{logo}}") {
          try{
            shape.replaceWithImage(logo, false);
          } catch (err){
            console.log(err.message);
            slide.insertImage(logo);
            shape.remove();
          }
          
        }
      });
    } else {
      // console.log("No logo");
      slide.getShapes().forEach(shape => {
        if (shape.getText().asString().trim() == "{{logo}}") {
          shape.remove();
        }
      });
    }

    // Promotion photo
    if (row[13] != ""){
      let regExp = new RegExp("[^=]*$");
      let photoId = regExp.exec(row[13])[0];

      var photo = UrlFetchApp.fetch(`https://drive.google.com/thumbnail?sz=w1000&id=${photoId}`, { headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() } }).getBlob(); // The endpoint is from https://stackoverflow.com/a/31504086

      // slide.insertImage(image);
      slide.getShapes().forEach(shape => {
        if (shape.getText().asString().trim() == "{{photo}}") {
          try{
            shape.replaceWithImage(photo, false);
          } catch (err){
            console.log(err.message);
            slide.insertImage(photo);
            shape.remove();
          }
        }
      });
    } else {
      console.log("No photo");
      slide.getShapes().forEach(shape => {
        if (shape.getText().asString().trim() == "{{photo}}") {
          shape.remove();
        }
      });
    }
  });

  // remove second slide (the template)
  // slides[1].remove();
  categories.forEach((category) => templateArr[category].remove());

  // index page
  arr.reverse();

  const indexPage = masterDeck.insertSlide(1);
  let textbox = indexPage.insertShape(SlidesApp.ShapeType.TEXT_BOX, 100, 200, 300, 60);
  let textRange = textbox.getText();
  arr.forEach((obj) => {
    let text = textRange.appendText(obj.name + "\n"); // circle name
    text.getTextStyle().setLinkSlide(obj.slideRef);
  });
  textRange.getTextStyle().setForegroundColor("#2e2e67");
}