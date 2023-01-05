import * as React from 'react';
import styles from './TransmittalApproveDocument.module.scss';
import { ITransmittalApproveDocumentProps } from './ITransmittalApproveDocumentProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class TransmittalApproveDocument extends React.Component<ITransmittalApproveDocumentProps, {}> {
  public render(): React.ReactElement<ITransmittalApproveDocumentProps> {

    return (
      <section className={`${styles.transmittalApproveDocument}`}>

      </section>
    );
  }
}
