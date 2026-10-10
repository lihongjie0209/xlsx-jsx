import { Workbook, Worksheet, Row, Cell, Image, Chart, Series } from "xlsx-jsx";

const logo = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAAAAAA6fptVAAAACklEQVR4nGNgAAAAAgABSK+kcQAAAABJRU5ErkJggg==";

export default <Workbook>
  <Worksheet name="Data">
    <Row><Cell>Category</Cell><Cell>Value</Cell></Row>
    <Row><Cell>One</Cell><Cell value={10} /></Row>
    <Row><Cell>Two</Cell><Cell value={20} /></Row>
  </Worksheet>
  <Worksheet name="Illustration">
    <Image cell="A1" width={96} height={48} data={logo} altText="Tiny sample logo" decorative={true} />
    <Chart type="column" cell="D2" width={240} height={120}>
      <Series values="Data!$B$2:$B$3" categories="Data!$A$2:$A$3" />
    </Chart>
    <Image cell="M1" width={48} height={24} data={logo} />
  </Worksheet>
</Workbook>;
