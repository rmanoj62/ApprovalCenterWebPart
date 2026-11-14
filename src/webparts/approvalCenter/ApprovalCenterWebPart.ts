import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';

import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import ApprovalCenter from './components/ApprovalCenter';

export interface IApprovalCenterWebPartProps {
}

export default class ApprovalCenterWebPart
  extends BaseClientSideWebPart<IApprovalCenterWebPartProps> {

  public render(): void {

    const element = React.createElement(
      ApprovalCenter,
      {
        siteUrl: this.context.pageContext.web.absoluteUrl,
        currentUserEmail: this.context.pageContext.user.email,
        spHttpClient: this.context.spHttpClient
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