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
  Created: string,
  CreatedPerson: string,
  Start: string,
  End: string,
  SubmissionStatus: string,
  Status: string,
  Categories: string,
  CategoriesNotSure?: string,
  RegistrationMethodology?: string,
  EmployeeCode?: string,
  EmployeeName?: string,
  AssistanceTypes?: string,
  Otherspecifyassistance?: string,
  Feedback?: string,
  Breaches?: string,
  Negative?: string,
  Name?: string,
  FirstName?: string,
  SecondName?: string,
  ThirdName?: string,
  LastName?: string,
  Sex?: string,
  Age?: string,
  Where?: string,
  Governorate?: string,
  SettlementGaza?: string,
  SettlementWestBank?: string,
  Assistance?: string,
  OngoingActivities?: string,
  OngoingActivitiesOther?: string,
  SubCategories?: string,
  SubCategoriesOther?: string,
  WhatIsYourId?: string,
  WhatIsYourPhoneNumber?: string,
  Feedback001?: string,
  Answer?: string,
  AnswerMode?: string,
  ModeOther?: string,
  ContactDetails?: string,
  Channels?: string,
  Specify?: string,
  HHsize?: string,
  DoYouWantToAssignToFocalPoint?: string,
  Sector?: string,
  SendEmailToFocalPoint?: string,
  FocalPoint?: string,
  FocalPointEmail?: string,
  CreatedBy: string,
  Modified: string,
  ModifiedBy: string
) {
  const values = [
    ID,
    Created,
    CreatedPerson,
    Start,
    End,
    SubmissionStatus,
    Status,
    Categories,
    CategoriesNotSure,
    RegistrationMethodology,
    EmployeeCode,
    EmployeeName,
    AssistanceTypes,
    Otherspecifyassistance,
    Feedback,
    Breaches,
    Negative,
    Name,
    FirstName,
    SecondName,
    ThirdName,
    LastName,
    Sex,
    Age,
    Where,
    Governorate,
    SettlementGaza,
    SettlementWestBank,
    Assistance,
    OngoingActivities,
    OngoingActivitiesOther,
    SubCategories,
    SubCategoriesOther,
    WhatIsYourId,
    WhatIsYourPhoneNumber,
    Feedback001,
    Answer,
    AnswerMode,
    ModeOther,
    ContactDetails,
    Channels,
    Specify,
    HHsize,
    DoYouWantToAssignToFocalPoint,
    Sector,
    SendEmailToFocalPoint,
    FocalPoint,
    CreatedBy,
    Modified,
    ModifiedBy,
  ];
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
