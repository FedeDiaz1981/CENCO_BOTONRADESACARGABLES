import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-webpart-base';

// 🔸 Inicializa los íconos de Fluent UI una sola vez al iniciar el webpart
// (podés dejar esta import; si tu build lo requiere, puedes cambiar a '@fluentui/font-icons-mdl2')
import { initializeIcons } from '@fluentui/react/lib/Icons';

import WpBotoneraDescargables from './components/WpBotoneraDescargables';
import { IWpBotoneraDescargablesProps } from './components/IWpBotoneraDescargablesProps';

export interface IWpBotoneraDescargablesWebPartProps {
  applyFilter: boolean;
  filterFieldInternalName: string;
  filterValue: string;
}

export default class WpBotoneraDescargablesWebPart extends BaseClientSideWebPart<IWpBotoneraDescargablesWebPartProps> {

  public onInit(): Promise<void> {
    initializeIcons();
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IWpBotoneraDescargablesProps> = React.createElement(
      WpBotoneraDescargables,
      {
        siteUrl: this.context.pageContext.site.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        applyFilter: this.properties.applyFilter,
        filterFieldInternalName: this.properties.filterFieldInternalName,
        filterValue: this.properties.filterValue
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Configuración' },
          groups: [
            {
              groupName: 'Filtro',
              groupFields: [
                PropertyPaneToggle('applyFilter', {
                  label: 'Aplicar filtro',
                  onText: 'Sí',
                  offText: 'No'
                }),
                PropertyPaneTextField('filterFieldInternalName', {
                  label: 'Campo de filtro (InternalName)'
                }),
                PropertyPaneTextField('filterValue', {
                  label: 'Valor a filtrar'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
