/* global AdminDirectory, DocumentApp */

// https://developers.google.com/apps-script/advanced/admin-sdk-directory

"use strict";

// eslint-disable-next-line no-unused-vars
function updateContactDoc() {
  const body = DocumentApp.getActiveDocument().getBody();
  let rangeElement;
  const managerName = AdminDirectory.Users
    .get("manager@sutherlandpipeband.org").name.fullName;
  const pmName = AdminDirectory.Users
    .get("pm@sutherlandpipeband.org").name.fullName;

  body.clear();

  body.editAsText()
    .appendText(
      `Please email Band Manager ${managerName} for general information or for booking:\n\n`
    ).appendText("manager@sutherlandpipeband.org\n\n")
    .appendText(
      `For questions about joining the band or learning to play, please contact Pipe\nMajor ${pmName}:\n\n`
    ).appendText("pm@sutherlandpipeband.org");

  // https://developers.google.com/apps-script/reference/document/body#findtextsearchpattern
  rangeElement = body.findText("manager@sutherlandpipeband.org");
  rangeElement.getElement()
    .asText()
    .setLinkUrl(
      rangeElement.getStartOffset(),
      rangeElement.getEndOffsetInclusive(),
      `mailto:${managerName}<manager@sutherlandpipeband.org>`
    );

  rangeElement = body.findText("pm@sutherlandpipeband.org");
  rangeElement.getElement()
    .asText()
    .setLinkUrl(
      rangeElement.getStartOffset(),
      rangeElement.getEndOffsetInclusive(),
      `mailto:${pmName}<pm@sutherlandpipeband.org>`
    );

  body.editAsText().setFontFamily(0, (body.getText().length - 1), "Arial");
  body.editAsText().setFontSize(0, (body.getText().length - 1), 12);
}
