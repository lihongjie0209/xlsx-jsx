import { Workbook, Worksheet, Row, Cell, Chart, Series } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Data">
    <Chart type="column" cell="F2" width={480} height={288}>
      <Series values="Data!$A$2:$A$3" fillColor="#4472C4" fillTransparency={10} />
      <Series values="Data!$B$2:$B$3" fillPattern={{ pattern: "diagBrick", foregroundColor: "#C00000", backgroundColor: "#FFF2CC" }} />
      <Series values="Data!$C$2:$C$3" fillGradient={{ type: "linear", angle: 45, stops: [{ color: "#4472C4", position: 0 }, { color: "#DDEBF7", position: 100 }] }} />
      <Series values="Data!$D$2:$D$3" fillHidden={true} lineColor="#17365D" lineWidth={1.25} />
    </Chart>
    <Row><Cell>Solid</Cell><Cell>Pattern</Cell><Cell>Gradient</Cell><Cell>Hidden</Cell></Row>
    <Row><Cell value={10} /><Cell value={12} /><Cell value={14} /><Cell value={16} /></Row>
    <Row><Cell value={20} /><Cell value={18} /><Cell value={22} /><Cell value={24} /></Row>
  </Worksheet>
</Workbook>;
