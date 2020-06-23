import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import UserPresence from './components/UserPresence';
import { IUserPresenceProps } from './components/IUserPresenceProps';

export default class UserPresenceWebPart extends BaseClientSideWebPart <{}> {
  public render(): void {
    const element: React.ReactElement<IUserPresenceProps> = React.createElement(
      UserPresence,
      {
        context: this.context
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
}
