function doGet(e) {
  const templateDocId = e.parameter.templateDocId;
  if(templateDocId == NaN || String(templateDocId) == "") {
    return createResponse(400, "Bad templateDocId");
  }
  
  return createResponse(200, getDocumentStructure(templateDocId));
}

function doPost(e) {
  try {
    var contents = JSON.parse(e.postData.contents);
  } catch (error) {
    return createResponse(400, "Bad JSON format");
  }

  var type = contents.type;
  var responseType = contents.responseType;
  var responseDocType = contents.responseDocType;

  if (type == "createDocFromTemplate") {
    docNew = createDocFromTemplate(
      contents.templateDocId,
      contents.folderId,
      contents.replacements
    );
  } else if (type == "mergeDocFromTemplate") {
    docNew = mergeDocFromTemplate(
      contents.templateDocId,
      contents.folderId,
      contents.mergeParameters
    );
  } else {
    return createResponse(400, "Bad type format", responseType);
  }

  if (responseDocType == "PDF") {
    response = docToBase64PDF(docNew);
  } else {
    response = docNew.getId();
  }

  return createResponse(200, response, responseType);
}

function createDocFromTemplate(templateDocId, FolderId, arrReplacements) {
  let docNew = makeCopyDocFileFromTemplate(templateDocId, FolderId);

  replacesInElement(docNew.getBody(), arrReplacements);

  return docNew;
}

function mergeDocFromTemplate(templateDocId, FolderId, arrMergeParameters) {
  let docNew = createDocFileFromTemplate(templateDocId, FolderId);

  mergeElementsFromTemplate(
    docNew.getBody(),
    DocumentApp.openById(templateDocId).getBody(),
    arrMergeParameters
  );

  return docNew;
}

function mergeElementsFromTemplate(
  newDocumentBody,
  templateDocBody,
  mergeParameters
) {
  var previousType = "";
  var previousElement = "";

  for (let i = 0; i < mergeParameters.length; i++) {
    templateElement = templateDocBody.getChild(mergeParameters[i].index).copy();
    type = templateElement.getType();

    if (type == DocumentApp.ElementType.PARAGRAPH) {
      newElement = newDocumentBody.appendParagraph(templateElement);
      previousElement = newElement;
      previousType = templateElement.getType();
    } else if (type == DocumentApp.ElementType.TABLE) {
      if (previousType == DocumentApp.ElementType.TABLE) {
        let row = templateElement.getRow(0).copy();
        newElement = previousElement.appendTableRow(row);
      } else {
        newElement = newDocumentBody.appendTable(templateElement);
        previousElement = newElement;
        previousType = type;
      }
    } else {
      continue;
    }

    replacesInElement(newElement, mergeParameters[i].replacements);
  }
}

function replacesInElement(element, arrReplacements) {
  arrReplacements.forEach((replacement) => {
    if (replacement.type == "text") {
      replaceTextToTextInElement(
        element,
        replacement.searchPattern,
        replacement.text
      );
    } else if (replacement.type == "image") {
      replaceTextToImageInElement(
        element,
        replacement.searchPattern,
        replacement.image,
        replacement.imageType,
        replacement.height,
        replacement.width
      );
    }
  });
}

function replaceTextToTextInElement(element, searchPattern, replacement) {
  element.replaceText(searchPattern, replacement);
}

function replaceTextToImageInElement(
  element,
  searchPattern,
  imageBase64,
  imageType,
  imageHeight = 50,
  imageWidth = 50
) {
  var rangeElement = element.findText(searchPattern);
  parentRangeElement = rangeElement.getElement().getParent();

  let image = parentRangeElement.addPositionedImage(
    base64ToBlob(imageBase64, imageType)
  );

  heightCurrent = image.getHeight();
  widthCurrent = image.getWidth();

  if (widthCurrent > imageWidth) {
    coefficient = imageWidth / widthCurrent;
    heightCurrent = Math.floor(heightCurrent * coefficient);
    widthCurrent = imageWidth;
  }

  if (heightCurrent > imageHeight) {
    coefficient = imageHeight / heightCurrent;
    widthCurrent = Math.floor(widthCurrent * coefficient);
    heightCurrent = imageHeight;
  }

  image.setHeight(heightCurrent);
  image.setWidth(widthCurrent);

  replaceTextToTextInElement(element, searchPattern, "");
}

function getDocumentStructure(docId) {
  var body = DocumentApp.openById(docId).getBody();

  response = [];

  var elements = body.getNumChildren();

  for (var i = 0; i < elements; i++) {
    var element = body.getChild(i);
    var text = element.getText();

    var pattern = /\{(|\/)v8 (.+?)\}/gm;

    response.push({
      index: i,
      match: text.match(pattern),
      type: element.getType(),
    });
  }

  return response;
}

function createDocFileFromTemplate(templateFileId, folderId) {
  let templateFile = DriveApp.getFileById(templateFileId);

  let document = DocumentApp.create(createFileName(templateFile));

  moveFileToFolder(document.getId(), folderId);

  return document;
}

function makeCopyDocFileFromTemplate(templateFileId, folderId) {
  const templateFile = DriveApp.getFileById(templateFileId);
  file = templateFile.makeCopy(createFileName(templateFile));

  moveFileToFolder(file.getId(), folderId);

  document = DocumentApp.openById(file.getId());

  return document;
}

function createFileName(templateFile) {
  templateFileName = templateFile.getName();

  var date = new Date();
  fileName = templateFileName + " " + date.toISOString();
  return fileName;
}

function moveFileToFolder(fileId, folderId) {
  if (folderId == NaN || String(folderId) === "" || String(fileId) === "") return;
  
  var file = DriveApp.getFileById(fileId);
  folder = DriveApp.getFolderById(folderId).addFile(file);
  file.getParents().next().removeFile(file);
}

function docToBase64PDF(doc) {
  var docBlob = doc.getAs("application/pdf");
  return Utilities.base64Encode(docBlob.getBytes());
}

function base64ToBlob(base64, contentType) {
  var decoded = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(decoded, contentType);

  return blob;
}

function createResponse(status = 400, data = "", responseType = "HTTP") {
  response = {
    status: status,
    data: data,
  };

  const stringResponse = JSON.stringify(response);

  if (responseType == "toAPI") {
    return stringResponse;
  } else {
    return ContentService.createTextOutput(stringResponse).setMimeType(
      ContentService.MimeType.JSON
    );
  }
}
