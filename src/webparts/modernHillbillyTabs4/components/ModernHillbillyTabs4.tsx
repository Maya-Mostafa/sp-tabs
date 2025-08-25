/* eslint-disable no-lone-blocks */
import * as React from 'react';
import styles from './ModernHillbillyTabs4.module.scss';
import type { IModernHillbillyTabs4Props } from './IModernHillbillyTabs4Props';
import { escape } from '@microsoft/sp-lodash-subset';
import { Tab, Tabs, TabList, TabPanel } from 'react-tabs';
import 'react-tabs/style/react-tabs.css';
import {  DisplayMode } from '@microsoft/sp-core-library';

export default class ModernHillbillyTabs4 extends React.Component<IModernHillbillyTabs4Props> {
  public render(): React.ReactElement<IModernHillbillyTabs4Props> {
    const {
      description,
      isDarkTheme,
      // environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    console.log("hillbilly props", this.props)

    const moveElement = (wpID: string) => {
      const el = document.getElementById(wpID);
      if (el) {
        const containerRef = (node: HTMLDivElement | null) => {
          if (node && el.parentNode !== node) {
            node.appendChild(el);
          }
        };
        return containerRef;
      }
    };

    const wpTabs = this.props.tabData ? Object.values(
      this.props.tabData.reduce((acc, { TabLabel, WebPartID }) => {
        if (!acc[TabLabel]) {
          acc[TabLabel] = { TabLabel, WebPartsIDs: [], WebPartsRefs: [] };
        }
        acc[TabLabel].WebPartsIDs.push(WebPartID);
        acc[TabLabel].WebPartsRefs.push(moveElement(WebPartID));
        return acc;
      }, {})
    ): [];

    console.log("wpTabs", wpTabs);

    return (
      <section className={`${styles.modernHillbillyTabs4} ${hasTeamsContext ? styles.teams : ''}`}>
        {this.props.displayMode === DisplayMode.Read ?
          <>
            { this.props.tabData && this.props.tabData.length === 0 ?
              <div className={styles.welcome}>
                <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
                {/* <h2>Well done, {escape(userDisplayName)}!</h2> */}
                {/* <div>{environmentMessage}</div> */}
                <div>Web part property value: <strong>{escape(description)}</strong></div>
              </div>
              :
              <div> 
                <Tabs forceRenderTabPanel={true}>
                  <TabList>
                    {wpTabs.map((tab: any, index) => (
                      <Tab key={index}>{tab.TabLabel}</Tab>
                    ))}
                  </TabList>
                  {wpTabs.map((tab: any, index) => (
                    <TabPanel key={index}>
                      {tab.WebPartsRefs.map((wpRef: string, wpIndex: number) => (
                        <div key={wpIndex} ref={wpRef} />
                      ))}
                    </TabPanel>
                  ))}
                </Tabs>          
              </div>
            }
          </>
          :
          <div className={styles.welcome}>
            <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
            <h2>You are in edit mode, {escape(userDisplayName)}!</h2>
            {/* <div>{environmentMessage}</div>
            <div>Web part property value: <strong>{escape(description)}</strong></div> */}
            {wpTabs.map((wp: any, wpIndex: number)=> (
              <div key={wpIndex}>Tab: {wp.TabLabel}</div>
            ))}
          </div>
        }
      </section>
    );
  }
}
