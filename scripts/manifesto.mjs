// Gera os manifestos do suplemento: manifest.xml (GitHub Pages) e manifest-dev.xml (https://localhost:3000).
// Uso: node scripts/manifesto.mjs
import { writeFileSync } from "node:fs";

const VERSAO = "1.0.0.0";
const alvos = [
  { arquivo: "manifest.xml", id: "91a11d27-49d3-441e-b6c0-e4b5ef6543d0", base: "https://rafaelmatos81.github.io/cronograma-pptx/", nome: "Cronograma" },
  { arquivo: "manifest-dev.xml", id: "91a11d27-49d3-441e-b6c0-e4b5ef6543d1", base: "https://localhost:3000/", nome: "Cronograma (dev)" },
];

const manifesto = ({ id, base, nome }) => `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="TaskPaneApp">
  <Id>${id}</Id>
  <Version>${VERSAO}</Version>
  <ProviderName>Rafael Matos</ProviderName>
  <DefaultLocale>pt-BR</DefaultLocale>
  <DisplayName DefaultValue="${nome}" />
  <Description DefaultValue="Gera Gantt e linha do tempo editáveis a partir de Excel, CSV ou MS Project XML, com linha de base, desvio, caminho crítico e faixas." />
  <IconUrl DefaultValue="${base}icones/icone-32.png" />
  <HighResolutionIconUrl DefaultValue="${base}icones/icone-64.png" />
  <SupportUrl DefaultValue="${base}" />
  <AppDomains>
    <AppDomain>${new URL(base).origin}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Presentation" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.4">
      <Set Name="PowerPointApi" MinVersion="1.4" />
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="${base}index.html" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Cronograma.Grupo">
                <Label resid="Cronograma.Grupo.Rotulo" />
                <Icon>
                  <bt:Image size="16" resid="Icone.16" />
                  <bt:Image size="32" resid="Icone.32" />
                  <bt:Image size="80" resid="Icone.80" />
                </Icon>
                <Control xsi:type="Button" id="Cronograma.Abrir">
                  <Label resid="Cronograma.Botao.Rotulo" />
                  <Supertip>
                    <Title resid="Cronograma.Botao.Rotulo" />
                    <Description resid="Cronograma.Botao.Dica" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icone.16" />
                    <bt:Image size="32" resid="Icone.32" />
                    <bt:Image size="80" resid="Icone.80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Cronograma.Painel</TaskpaneId>
                    <SourceLocation resid="Painel.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icone.16" DefaultValue="${base}icones/icone-16.png" />
        <bt:Image id="Icone.32" DefaultValue="${base}icones/icone-32.png" />
        <bt:Image id="Icone.80" DefaultValue="${base}icones/icone-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Painel.Url" DefaultValue="${base}index.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Cronograma.Grupo.Rotulo" DefaultValue="${nome}" />
        <bt:String id="Cronograma.Botao.Rotulo" DefaultValue="${nome}" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Cronograma.Botao.Dica" DefaultValue="Abre o painel para criar ou editar um cronograma (Gantt) a partir de Excel, CSV ou MS Project." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;

for (const a of alvos) {
  writeFileSync(a.arquivo, manifesto(a));
  console.log(`${a.arquivo} -> ${a.base}`);
}
