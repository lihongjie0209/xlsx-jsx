import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Font properties">
    <Row>
      <Cell fontName="Aptos Display" fontSize={12.5} fontFamily={3} fontCharset={128} fontScheme="headings">Custom font family and size</Cell>
      <Cell value={42} fontName="Arial" fontSize={10} bold={true} />
      <Cell formula="=1+1" fontSize={14} fontScheme="none" numberFormat="0.00" />
    </Row>
  </Worksheet>
</Workbook>;
