import * as React from 'react';
import * as ReactDom from 'react-dom';

import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { graph } from '@pnp/graph';
import Users from './components/Users';
import { IUsersProps } from './components/IUsersProps';

export interface IUsersWebPartProps { }

export default class AzureADUsersWebPart extends BaseClientSideWebPart<IUsersWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      // Initialize PnP
      graph.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IUsersProps> = React.createElement(
      Users, { }
    );

    this.context.statusRenderer.displayLoadingIndicator(this.domElement, "Loading users");

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
