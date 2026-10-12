import { Workbook, Worksheet, Row, Cell, Chart, Series } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Data">
    <Row><Cell>Category</Cell><Cell>Value</Cell></Row>
    <Row><Cell>North</Cell><Cell value={18} /></Row>
    <Row><Cell>South</Cell><Cell value={27} /></Row>
    <Row><Cell>East</Cell><Cell value={34} /></Row>
    <Row><Cell>West</Cell><Cell value={22} /></Row>
  </Worksheet>
  <Worksheet name="Chart">
    <Chart type="column" cell="D2">
      <Series
        values="Data!$B$2:$B$5"
        categories="Data!$A$2:$A$5"
        pointColors={{1: "#336699", 3: "#F0A020"}}
        pointColorTransparencies={{1: 25}}
        pointLines={{1: {color: "#17365D", width: 1.25, dash: "dash"}, 2: {hidden: true}}}
        pointPatterns={{2: {pattern: "pct20", foregroundColor: "#C00000", backgroundColor: "#FFF2CC"}}}
      />
    </Chart>
  </Worksheet>
</Workbook>;
