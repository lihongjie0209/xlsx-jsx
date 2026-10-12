import { Workbook, Worksheet, Row, Cell, Chart, Series } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Data">
    <Row><Cell>Category</Cell><Cell>Value</Cell></Row>
    <Row><Cell>One</Cell><Cell value={10} /></Row>
    <Row><Cell>Two</Cell><Cell value={20} /></Row>
    <Row><Cell>Three</Cell><Cell value={30} /></Row>
  </Worksheet>
  <Worksheet name="Chart">
    <Chart type="column" cell="D2" titleFormula="Data!$A$1">
      <Series
        values="Data!$B$2:$B$4"
        categories="Data!$A$2:$A$4"
        customLabels={["Forecast", null, null]}
        pointLabels={{
          0: { fontSize: 14, fontBold: true, fontColor: "#123456", fillColor: "#FFF2CC" },
          2: { text: "Peak", numberFormat: "0.0", lineColor: "#654321", lineWidth: 1 }
        }}
      />
    </Chart>
  </Worksheet>
</Workbook>;
