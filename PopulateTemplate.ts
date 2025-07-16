/* 
{
  Id
  Title
  Created
  CreatedPerson
  Start
  End
  SubmissionStatus
  Status
  Categories
  CategoriesNotSure
  RegistrationMethodology
  EmployeeCode
  EmployeeName
  AssistanceTypes
  Otherspecifyassistance
  Feedback
  Breaches
  Negative
  Name
  FirstName
  SecondName
  ThirdName
  LastName
  Sex
  Age
  Where
  Governorate
  SettlementGaza
  SettlementWestBank
  Assistance
  OngoingActivities
  OngoingActivitiesOther
  SubCategories
  SubCategoriesOther
  WhatIsYourId
  WhatIsYourPhoneNumber
  Feedback001
  Answer
  AnswerMode
  ModeOther
  ContactDetails
  Channels
  Specify
  Hhsize
  DoYouWantToAssignToFocalPoint
  Sector
  SendEmailToFocalPoint
  FocalPoint
  FocalPointEmail
  CreatedBy
  Modified
  ModifiedBy
}
*/

function main(
  workbook: ExcelScript.Workbook,
  ID: number,
  Title: string,
  Created: string
) {
  const values = [ID, Title, Created];
  const sheet = workbook.getWorksheet("Update oPt Cases");

  // Always start checking from row 2 (index 1)
  const startRow = 1; // Row 2 in Excel (0-based index)

  // Read a reasonable number of rows in column A
  const colA = sheet.getRange(`A${startRow + 1}:A1000`).getValues();

  let writeRow = startRow;
  for (let i = 0; i < colA.length; i++) {
    if (colA[i][0] === "") {
      writeRow = startRow + i;
      break;
    }
  }

  // Write to the first blank row starting from row 2
  sheet.getRangeByIndexes(writeRow, 0, 1, values.length).setValues([values]);
}
