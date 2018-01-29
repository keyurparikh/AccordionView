import * as React from 'react';
import styles from './AccordionView.module.scss';
import { IAccordionViewProps } from './IAccordionViewProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class AccordionView extends React.Component<IAccordionViewProps, {}> {
  public render(): React.ReactElement<IAccordionViewProps> {
    return (
      <div className="accordion">
         <div className={ styles.accordionView }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Accordion WebPart!</span>
              <p className={ styles.subTitle }>Edit web part property and select Accordion list.</p>
              <p className={ styles.description }>Accordion list must have Title and Body fields.</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
      </div>
    );
  }
}
