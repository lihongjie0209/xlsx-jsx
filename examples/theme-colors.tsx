import { Workbook, Worksheet, Row, Cell, Column } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Theme colors">
    <Column min={1} max={1} fontColor={{ theme: 4, shade: 2 }} />
    <Row fillPattern="darkDown"
      fillForegroundColor={{ theme: 5, shade: 1 }}
      fillBackgroundColor={{ theme: 0, shade: 3 }}>
      <Cell value="Theme foreground/background"
        fontColor={{ theme: 2, shade: 0 }}
        fillColor={{ theme: 6, shade: 3 }}
        borderBottom="thin"
        borderBottomColor={{ theme: 1, shade: 4 }} />
    </Row>
  </Worksheet>
</Workbook>;
