/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable no-lone-blocks */
import * as React from 'react';
import styles from './ModernHillbillyTabs4.module.scss';
import './ModernHillbillyTabs4.scss';
import type { IModernHillbillyTabs4Props } from './IModernHillbillyTabs4Props';
import { escape } from '@microsoft/sp-lodash-subset';
import { Tab, Tabs, TabList, TabPanel } from 'react-tabs';
import 'react-tabs/style/react-tabs.css';
import {  DisplayMode } from '@microsoft/sp-core-library';
import { Icon } from '@fluentui/react';

// export default class ModernHillbillyTabs4 extends React.Component<IModernHillbillyTabs4Props> {
  // public render(): React.ReactElement<IModernHillbillyTabs4Props> {

export default function ModernHillbillyTabs4 (props:IModernHillbillyTabs4Props){
    const {
      description,
      isDarkTheme,
      // environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = props;

    const [tabIndex, setTabIndex] = React.useState(0);

   
const queryParams = new URLSearchParams(window.location.search);
        const tabParam = Number(queryParams.get("tab"));
    // console.log("tabParam", tabParam)
    const tabSelectHandler = (index: number) => {
      setTabIndex(index);
      if (queryParams.has('tab')) queryParams.delete('tab');
      window.history.replaceState({}, '', `?tab=${index}`);
    };

    React.useEffect(()=>{
       
        if (queryParams.has("tab")){
          setTabIndex(tabParam);
        }
    }, [tabIndex]);

    console.log("hillbilly props", props);
    console.log("tabIndex", tabIndex);

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

    const wpTabs = props.tabData ? Object.values(
      props.tabData.reduce((acc, { TabLabel, WebPartID, TabIcon, TabBgColor, TabForColor }) => {
        if (!acc[TabLabel]) {
          acc[TabLabel] = { TabLabel, WebPartsIDs: [], WebPartsRefs: [], TabIcon: '', TabBgColor: '', TabForColor:'' };
        }
        acc[TabLabel].WebPartsIDs.push(WebPartID);
        acc[TabLabel].WebPartsRefs.push(moveElement(WebPartID));
        acc[TabLabel].TabIcon = TabIcon;
        acc[TabLabel].TabBgColor = TabBgColor ? TabBgColor.str : '';
        acc[TabLabel].TabForColor = TabForColor ? TabForColor.str : '';
        return acc;
      }, {})
    ): [];

    console.log("wpTabs", wpTabs);

    return (
      <section className={`${styles.modernHillbillyTabs4} ${hasTeamsContext ? styles.teams : ''} ${props.tabStyle} ${props.tabAlign}`}>
        {props.displayMode === DisplayMode.Read ?
          <>
            { props.tabData && props.tabData.length === 0 ?
              <div className={styles.welcome}>
                <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
                {/* <h2>Well done, {escape(userDisplayName)}!</h2> */}
                {/* <div>{environmentMessage}</div> */}
                <div>Web part property value: <strong>{escape(description)}</strong></div>
              </div>
              :
              <div> 
                <Tabs 
                  forceRenderTabPanel={true} 
                  selectedIndex={tabIndex} 
                  onSelect={tabSelectHandler}
                >
                  <TabList>
                    {wpTabs.map((tab: any, index) => (
                      <Tab key={index} style={{backgroundColor: tab.TabBgColor, color: tab.TabForColor}}>
                        <Icon style={{color: tab.TabForColor}} className='tabIcon' iconName={tab.TabIcon} />
                        <div style={{color: tab.TabForColor}}>{tab.TabLabel}</div>
                        <div className='activeTabTriangle'/>
                      </Tab>
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
            <Tabs 
              forceRenderTabPanel={true} 
              selectedIndex={tabIndex} 
              onSelect={tabSelectHandler}
            >
              <TabList>
                {wpTabs.map((tab: any, index) => (
                  <Tab key={index} style={{backgroundColor: tab.TabBgColor, color: tab.TabForColor}}>
                    <Icon style={{color: tab.TabForColor}} className='tabIcon' iconName={tab.TabIcon} />
                    <div style={{color: tab.TabForColor}}>{tab.TabLabel}</div>
                    <div className='activeTabTriangle'/>
                  </Tab>
                ))}
              </TabList>
              {wpTabs.map((wp: any, wpIndex: number)=> (
                <div key={wpIndex}>Tab: {wp.TabLabel}</div>
              ))}
            </Tabs>     

          </div>
        }
      </section>
    );
  }

