import * as React from 'react';
import styles from './CustomPropertyPaneWp.module.scss';
import { ICustomPropertyPaneWpProps } from './ICustomPropertyPaneWpProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class CustomPropertyPaneWp extends React.Component<ICustomPropertyPaneWpProps, {}> {
  public render(): React.ReactElement<ICustomPropertyPaneWpProps> {
    return (
      <div className={ styles.customPropertyPaneWp }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.listName)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
