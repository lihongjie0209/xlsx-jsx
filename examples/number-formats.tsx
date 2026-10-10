import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Formats">
    <Row>
      <Cell value={1234.5} numberFormat={"#,##0.00"} />
      <Cell value={0.125} numberFormat={"0.00%"} />
      <Cell formula={"=A1*2"} numberFormat={"$#,##0.00"} />
      <Cell date={"2026-10-10"} numberFormat={"dd/mm/yyyy"} />
      <Cell value={1.5} numberFormat={"[h]:mm:ss"} />
    </Row>
  </Worksheet>
</Workbook>;
