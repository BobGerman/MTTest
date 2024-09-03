import * as React from 'react';
import styles from './SpfxHelloWorld.module.scss';
import type { ISpfxHelloWorldProps } from './ISpfxHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SpfxHelloWorld extends React.Component<ISpfxHelloWorldProps, {}> {
  public render(): React.ReactElement<ISpfxHelloWorldProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      loginName
    } = this.props;

    return (
      <section className={`${styles.spfxHelloWorld} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            Hello {this.props.userDisplayName} ({loginName})!
          </p>
          <h3>REST call results</h3>
          <ol>
            {this.props.results?.length > 0 && this.props.results.map((r, idx) => {
              return (<li key={idx} >{r}</li>);
            })}
            {(!this.props.results || this.props.results.length < 1) &&
              <p>No results to report!</p>
            }
          </ol>
        </div>
      </section>
    );
  }
}
