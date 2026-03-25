import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import ProductChart from './components/ProductChart';

export interface IProductChartWebPartProps {
  description: string;
}

export default class ProductChartWebPart extends BaseClientSideWebPart<IProductChartWebPartProps> {


public render(): void {
  const element: React.ReactElement<any> = React.createElement(
    ProductChart,
    {
      context: this.context
    }
  );

  ReactDom.render(element, this.domElement);
}





}
