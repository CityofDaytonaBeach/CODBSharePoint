// ============================================================================
// Component Template Registry
// Pluggable, authorable templates for Web Parts, ACEs, and Extensions.
// The Designer uses these to scaffold new components instead of a fixed blob.
// ============================================================================

import type { ComponentDefinition, ExtensionDefinition, ACEDefinition } from '../types/index.js';

export type TemplateKind = 'webpart' | 'ace' | 'extension';

export interface TemplateRenderContext {
  component: ComponentDefinition;
  extension?: ExtensionDefinition;
  ace?: ACEDefinition;
  namespace: string;
}

export interface ComponentTemplate {
  name: string;
  kind: TemplateKind;
  framework?: 'react' | 'vanilla';
  description: string;
  render(ctx: TemplateRenderContext): Map<string, string>;
}

const DEFAULT_WEBPART_STYLE = (name: string) => `@import '~@microsoft/sp-tslint-theme/vars.scss';

.${name} {
  .container {
    max-width: 700px;
    margin: 0px auto;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
    padding: 20px;
  }
  .loading { text-align: center; padding: 20px; }
  .content { margin-top: 10px; }
  .item { padding: 10px; border-bottom: 1px solid #eee; }
}
`;

const REACT_WEBPART: ComponentTemplate = {
  name: 'react-webpart',
  kind: 'webpart',
  framework: 'react',
  description: 'React class-component web part with a property pane',
  render({ component, namespace }) {
    const c = component;
    const dir = `src/webparts/${c.name}`;
    const files = new Map<string, string>();

    files.set(`${dir}/${c.name}WebPart.ts`, `import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import ${c.name}Component from './components/${c.name}';
import { I${c.name}Props } from './components/I${c.name}Props';

export interface I${c.name}WebPartProps { description: string; }

export default class ${c.name}WebPart extends BaseClientSideWebPart<I${c.name}WebPartProps> {
  public render(): void {
    const element: React.ReactElement<I${c.name}Props> = React.createElement(
      ${c.name}Component,
      { description: this.properties.description }
    );
    ReactDOM.render(element, this.domElement);
  }
  protected onDispose(): void { ReactDOM.unmountComponentAtNode(this.domElement); }
  protected get dataVersion(): Version { return Version.parse('1.0'); }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{ header: { description: '${c.displayName} Settings' }, groups: [{ groupName: 'Basic', groupFields: [PropertyPaneTextField('description', { label: 'Description' })] }] }]
    };
  }
}
`);

    files.set(`${dir}/components/I${c.name}Props.ts`, `export interface I${c.name}Props {
  description: string;
}
`);

    files.set(`${dir}/components/${c.name}.tsx`, `import * as React from 'react';
import styles from './${c.name}.module.scss';
import { I${c.name}Props } from './I${c.name}Props';

export default class ${c.name}Component extends React.Component<I${c.name}Props, {}> {
  public render(): React.ReactElement<I${c.name}Props> {
    return (
      <div className={styles.${c.name}}>
        <div className={styles.container}>
          <h2>${c.displayName}</h2>
          <p>{this.props.description}</p>
        </div>
      </div>
    );
  }
}
`);

    files.set(`${dir}/components/${c.name}.module.scss`, DEFAULT_WEBPART_STYLE(c.name));
    return files;
  }
};

const VANILLA_WEBPART: ComponentTemplate = {
  name: 'vanilla-webpart',
  kind: 'webpart',
  framework: 'vanilla',
  description: 'Lightweight DOM web part without React',
  render({ component, namespace }) {
    const c = component;
    const dir = `src/webparts/${c.name}`;
    const files = new Map<string, string>();
    files.set(`${dir}/${c.name}WebPart.ts`, `import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

export default class ${c.name}WebPart extends BaseClientSideWebPart<{ description: string }> {
  public render(): void {
    this.domElement.innerHTML = \`
      <div class="${namespace}-${c.name}">
        <h2>${c.displayName}</h2>
        <p>\${this.properties.description}</p>
      </div>\`;
  }
  protected get dataVersion(): Version { return Version.parse('1.0'); }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ header: { description: '${c.displayName}' }, groups: [{ groupName: 'Settings', groupFields: [PropertyPaneTextField('description', { label: 'Description' })] }] }] };
  }
}
`);
    return files;
  }
};

const APP_CUSTOMIZER: ComponentTemplate = {
  name: 'application-customizer',
  kind: 'extension',
  description: 'Application Customizer extension',
  render({ extension }) {
    const e = extension as ExtensionDefinition;
    const dir = `src/extensions/${e.name}`;
    const files = new Map<string, string>();
    files.set(`${dir}/${e.name}ApplicationCustomizer.ts`, `import { Log } from '@microsoft/sp-core-library';
import { type IApplicationCustomizerProperties, BaseApplicationCustomizer } from '@microsoft/sp-application-base';

export default class ${e.name}ApplicationCustomizer extends BaseApplicationCustomizer<IApplicationCustomizerProperties> {
  public onInit(): Promise<void> { Log.info('${e.name}', 'Initialized'); return Promise.resolve(); }
}
`);
    return files;
  }
};

const FIELD_CUSTOMIZER: ComponentTemplate = {
  name: 'field-customizer',
  kind: 'extension',
  description: 'Field Customizer extension',
  render({ extension }) {
    const e = extension as ExtensionDefinition;
    const dir = `src/extensions/${e.name}`;
    const files = new Map<string, string>();
    files.set(`${dir}/${e.name}FieldCustomizer.ts`, `import { Log } from '@microsoft/sp-core-library';
import { type IFieldCustomizerCellProperties, BaseFieldCustomizer } from '@microsoft/sp-listview-extensibility';

export default class ${e.name}FieldCustomizer extends BaseFieldCustomizer<IFieldCustomizerCellProperties> {
  public onInit(): Promise<void> { return Promise.resolve(); }
  public onRenderCell(): void { /* render cell */ }
}
`);
    return files;
  }
};

const COMMAND_SET: ComponentTemplate = {
  name: 'command-set',
  kind: 'extension',
  description: 'ListView Command Set extension',
  render({ extension }) {
    const e = extension as ExtensionDefinition;
    const dir = `src/extensions/${e.name}`;
    const files = new Map<string, string>();
    files.set(`${dir}/${e.name}CommandSet.ts`, `import { Log } from '@microsoft/sp-core-library';
import { type IListViewCommandSetProperties, BaseListViewCommandSet, RowClickedEvent } from '@microsoft/sp-listview-extensibility';

export default class ${e.name}CommandSet extends BaseListViewCommandSet<IListViewCommandSetProperties> {
  public onInit(): Promise<void> { Log.info('${e.name}', 'Init'); return Promise.resolve(); }
  public onListViewSelectedChanged(event: RowClickedEvent): void { /* handle */ }
}
`);
    return files;
  }
};

const FORM_CUSTOMIZER: ComponentTemplate = {
  name: 'form-customizer',
  kind: 'extension',
  description: 'Form Customizer extension',
  render({ extension }) {
    const e = extension as ExtensionDefinition;
    const dir = `src/extensions/${e.name}`;
    const files = new Map<string, string>();
    files.set(`${dir}/${e.name}FormCustomizer.ts`, `import { Log } from '@microsoft/sp-core-library';
import { type IFormCustomizerProperties, BaseFormCustomizer } from '@microsoft/sp-listview-extensibility';

export default class ${e.name}FormCustomizer extends BaseFormCustomizer<IFormCustomizerProperties> {
  public onInit(): Promise<void> { return Promise.resolve(); }
  public onRender(): void { this.domElement.innerHTML = \`<div>${e.displayName}</div>\`; }
}
`);
    return files;
  }
};

const ACE_TEMPLATE: ComponentTemplate = {
  name: 'ace',
  kind: 'ace',
  description: 'Adaptive Card Extension for Viva Connections',
  render({ component, ace }) {
    const a = ace as ACEDefinition;
    const c = component;
    const dir = `src/quickView/${a.name}`;
    const files = new Map<string, string>();
    files.set(`${dir}/${a.name}Card.ts`, `import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { IQuickView } from '@microsoft/sp-adaptive-card-extension-base';

export default class ${a.name}ACE extends BaseAdaptiveCardExtension<{}, IQuickView> {
  public get title(): string { return '${c.displayName}'; }
  public renderCard(): string { return '${c.displayName}'; }
}
`);
    return files;
  }
};

const DEFAULT_TEMPLATES: ComponentTemplate[] = [
  REACT_WEBPART,
  VANILLA_WEBPART,
  APP_CUSTOMIZER,
  FIELD_CUSTOMIZER,
  COMMAND_SET,
  FORM_CUSTOMIZER,
  ACE_TEMPLATE
];

export class TemplateRegistry {
  private templates = new Map<string, ComponentTemplate>();

  constructor(templates: ComponentTemplate[] = DEFAULT_TEMPLATES) {
    for (const t of templates) this.register(t);
  }

  register(template: ComponentTemplate): void {
    this.templates.set(template.name, template);
  }

  unregister(name: string): boolean {
    return this.templates.delete(name);
  }

  get(name: string): ComponentTemplate | undefined {
    return this.templates.get(name);
  }

  list(): ComponentTemplate[] {
    return Array.from(this.templates.values());
  }

  resolveForComponent(component: ComponentDefinition, extension?: ExtensionDefinition, ace?: ACEDefinition): ComponentTemplate | undefined {
    if (component.type === 'ace') return this.templates.get('ace');
    if (component.type === 'extension' && extension) {
      switch (extension.type) {
        case 'ApplicationCustomizer': return this.templates.get('application-customizer');
        case 'FieldCustomizer': return this.templates.get('field-customizer');
        case 'ListViewCommandSet': return this.templates.get('command-set');
        case 'FormCustomizer': return this.templates.get('form-customizer');
        default: return undefined;
      }
    }
    if (component.framework === 'react') return this.templates.get('react-webpart');
    return this.templates.get('vanilla-webpart');
  }
}

export { DEFAULT_WEBPART_STYLE };
