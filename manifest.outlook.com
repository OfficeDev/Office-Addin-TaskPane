<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>05c2e1c9-3e1d-406e-9a91-e9ac64854143</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Contoso Task Pane Add-in"/>
  <Description DefaultValue="A template to get started."/>
  <IconUrl DefaultValue="https://localhost:3000/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://localhost:3000/assets/icon-80.png"/>
  <SupportUrl DefaultValue="https://www.contoso.com/help"/>
  <AppDomains>
    <AppDomain>contoso.com</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document"/>
    <Host Name="Mailbox" />
    <Host Name="Notebook"/>
    <Host Name="Presentation"/>
    <Host Name="Project"/>
    <Host Name="Workbook"/>
  </Hosts>  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
  </DefaultSettings>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="https://localhost:3000/index.html"/>
        <RequestedHeight>250</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.1" />
    </Sets>
  </Requirements>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="Contoso.GetStarted.Title"/>
            <Description resid="Contoso.GetStarted.Description"/>
            <LearnMoreUrl resid="Contoso.GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Contoso.Ribbon.Url" />
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Contoso.Group1">
                <Label resid="Contoso.Group1Label" />
                <Icon>
                  <bt:Image size="16" resid="Contoso.tpicon_16x16" />
                  <bt:Image size="32" resid="Contoso.tpicon_32x32" />
                  <bt:Image size="80" resid="Contoso.tpicon_80x80" />
                </Icon>
                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.tpicon_16x16" />
                    <bt:Image size="32" resid="Contoso.tpicon_32x32" />
                    <bt:Image size="80" resid="Contoso.tpicon_80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Contoso.Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="Contoso.Ribbon.Url" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadGroup">
                <Label resid="Contoso.Group1Label" />
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon16" />
                    <bt:Image size="32" resid="icon32" />
                    <bt:Image size="80" resid="icon80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Contoso.Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
      <Host xsi:type="Notebook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="Contoso.GetStarted.Title"/>
            <Description resid="Contoso.GetStarted.Description"/>
            <LearnMoreUrl resid="Contoso.GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Contoso.Ribbon.Url" />
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Contoso.Group1">
                <Label resid="Contoso.Group1Label" />
                <Icon>
                  <bt:Image size="16" resid="Contoso.tpicon_16x16" />
                  <bt:Image size="32" resid="Contoso.tpicon_32x32" />
                  <bt:Image size="80" resid="Contoso.tpicon_80x80" />
                </Icon>
                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.tpicon_16x16" />
                    <bt:Image size="32" resid="Contoso.tpicon_32x32" />
                    <bt:Image size="80" resid="Contoso.tpicon_80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Contoso.Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="Contoso.GetStarted.Title"/>
            <Description resid="Contoso.GetStarted.Description"/>
            <LearnMoreUrl resid="Contoso.GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Contoso.Ribbon.Url" />
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Contoso.Group1">
                <Label resid="Contoso.Group1Label" />
                <Icon>
                  <bt:Image size="16" resid="Contoso.tpicon_16x16" />
                  <bt:Image size="32" resid="Contoso.tpicon_32x32" />
                  <bt:Image size="80" resid="Contoso.tpicon_80x80" />
                </Icon>
                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.tpicon_16x16" />
                    <bt:Image size="32" resid="Contoso.tpicon_32x32" />
                    <bt:Image size="80" resid="Contoso.tpicon_80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Contoso.Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
      <Host xsi:type="Project">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="Contoso.GetStarted.Title"/>
            <Description resid="Contoso.GetStarted.Description"/>
            <LearnMoreUrl resid="Contoso.GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Contoso.Ribbon.Url" />
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Contoso.Group1">
                <Label resid="Contoso.Group1Label" />
                <Icon>
                  <bt:Image size="16" resid="Contoso.tpicon_16x16" />
                  <bt:Image size="32" resid="Contoso.tpicon_32x32" />
                  <bt:Image size="80" resid="Contoso.tpicon_80x80" />
                </Icon>
                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.tpicon_16x16" />
                    <bt:Image size="32" resid="Contoso.tpicon_32x32" />
                    <bt:Image size="80" resid="Contoso.tpicon_80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Contoso.Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="Contoso.GetStarted.Title"/>
            <Description resid="Contoso.GetStarted.Description"/>
            <LearnMoreUrl resid="Contoso.GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Contoso.Ribbon.Url" />
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Contoso.Group1">
                <Label resid="Contoso.Group1Label" />
                <Icon>
                  <bt:Image size="16" resid="Contoso.tpicon_16x16" />
                  <bt:Image size="32" resid="Contoso.tpicon_32x32" />
                  <bt:Image size="80" resid="Contoso.tpicon_80x80" />
                </Icon>
                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.tpicon_16x16" />
                    <bt:Image size="32" resid="Contoso.tpicon_32x32" />
                    <bt:Image size="80" resid="Contoso.tpicon_80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Contoso.Taskpane.Url" />
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
        <bt:Image id="Contoso.tpicon_16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Contoso.tpicon_32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Contoso.tpicon_80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Contoso.GetStarted.LearnMoreUrl" DefaultValue="https://go.microsoft.com/fwlink/?LinkId=276812" />
        <bt:Url id="Contoso.Ribbon.Url" DefaultValue="https://localhost:3000/ribbon.html" />
        <bt:Url id="Contoso.Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Contoso.GetStarted.Title" DefaultValue="Get started with your sample add-in!" />
        <bt:String id="Contoso.Group1Label" DefaultValue="Commands Group" />
        <bt:String id="Contoso.TaskpaneButton.Label" DefaultValue="Show Taskpane" />
        <bt:String id="customTabLabel"  DefaultValue="My Add-in Tab"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Contoso.GetStarted.Description" DefaultValue="Your sample add-in loaded succesfully. Go to the HOME tab and click the 'Show Taskpane' button to get started." />
        <bt:String id="Contoso.TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>