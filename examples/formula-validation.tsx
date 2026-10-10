import { Workbook, Worksheet, DataValidation } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Input">
      <DataValidation range="A2:A20" type="whole" formula1="1" formula2="100" />
      <DataValidation range="B2:B20" type="decimal" operator="greaterThanOrEqual" formula1="0" />
      <DataValidation range="C2:C20" type="date" operator="between" formula1="DATE(2026,1,1)" formula2="DATE(2026,12,31)" />
      <DataValidation range="D2:D20" type="time" operator="between" formula1="TIME(9,0,0)" formula2="TIME(17,0,0)" />
      <DataValidation range="E2:E20" type="textLength" operator="lessThanOrEqual" formula1="80" />
      <DataValidation range="F2:F20" type="custom" formula1="MOD(F2,2)=0" />
    </Worksheet>
  </Workbook>
);
