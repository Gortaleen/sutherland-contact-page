/**
 * https://github.com/Gortaleen/sutherland-contact-page
 * https://github.com/google/clasp#readme
 * https://github.com/google/clasp/blob/master/docs/typescript.md
 * https://khalilstemmler.com/blogs/typescript/eslint-for-typescript/
 * https://www.typescriptlang.org/docs/handbook/release-notes/typescript-2-0.html#non-null-assertion-operator
 * https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/Nullish_coalescing
 * https://www.typescriptlang.org/tsconfig#strict
 * https://www.typescriptlang.org/tsconfig#alwaysStrict
 * https://typescript-eslint.io/getting-started
 * https://prettier.io/docs/en/integrating-with-linters.html
 */

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function updateContactRun() {
  ContactUpdate.main();
}

const ContactUpdate = (function () {
  function main() {
    const docId =
      PropertiesService.getScriptProperties().getProperty("DOCUMENT_ID");
    if (!docId) {
      throw "Document ID is missing";
    }
    const body = DocumentApp.openById(docId).getBody();
    let rangeElement;
    if (!AdminDirectory.Users) {
      throw "AdminDirectory.Users.get not available";
    }
    const managerName = AdminDirectory.Users.get(
      "manager@sutherlandpipeband.org"
    ).name?.fullName;

    // https://developers.google.com/apps-script/advanced/admin-sdk-directory
    const pmName = AdminDirectory.Users.get("pm@sutherlandpipeband.org").name
      ?.fullName;

    body.clear();

    body
      .editAsText()
      .appendText(
        `Please email Band Manager ${managerName} for general information or for booking:\n\n`
      )
      .appendText("manager@sutherlandpipeband.org\n\n")
      .appendText(
        `For questions about joining the band or learning to play, please contact Pipe\nMajor ${pmName}:\n\n`
      )
      .appendText("pm@sutherlandpipeband.org");

    // https://developers.google.com/apps-script/reference/document/body#findtextsearchpattern
    rangeElement = body.findText("manager@sutherlandpipeband.org");
    rangeElement
      .getElement()
      .asText()
      .setLinkUrl(
        rangeElement.getStartOffset(),
        rangeElement.getEndOffsetInclusive(),
        `mailto:${managerName}<manager@sutherlandpipeband.org>`
      );

    rangeElement = body.findText("pm@sutherlandpipeband.org");
    rangeElement
      .getElement()
      .asText()
      .setLinkUrl(
        rangeElement.getStartOffset(),
        rangeElement.getEndOffsetInclusive(),
        `mailto:${pmName}<pm@sutherlandpipeband.org>`
      );

    body.editAsText().setFontFamily(0, body.getText().length - 1, "Arial");
    body.editAsText().setFontSize(0, body.getText().length - 1, 12);
  }

  return { main };
})();
