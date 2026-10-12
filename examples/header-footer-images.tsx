import { FooterImage, HeaderImage, Row, Cell, Workbook, Worksheet } from "xlsx-jsx";

const logo =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAAAAAA6fptVAAAACklEQVR4nGNgAAAAAgABSK+kcQAAAABJRU5ErkJggg==";

export default (
  <Workbook>
    <Worksheet name="HeaderFooter" header="&L&G" footer="&R&G">
      <HeaderImage position="left" data={logo} />
      <FooterImage position="right" data={logo} />
      <Row><Cell>VML header/footer image validation</Cell></Row>
    </Worksheet>
  </Workbook>
);
