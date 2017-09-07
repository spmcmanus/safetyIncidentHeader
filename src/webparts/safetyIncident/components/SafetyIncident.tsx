// React and jQuery libs
import * as React from 'react';
import * as jquery from 'jquery';

// Styling
import styles from './SafetyIncident.module.scss';

// Office-Ui Fabric Components
import {
  Persona,
  PersonaInitialsColor,
} from 'office-ui-fabric-react/lib/Persona';
import {
  Image,
  IImageProps,
  ImageFit
} from 'office-ui-fabric-react/lib/Image';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';

// custom components
import { ISafetyIncidentProps } from './ISafetyIncidentProps';

// set up local state
export interface localState {
  incidentId: number;
  jobNumber: string;
  jobLocation: string;
  incidentDate: string;
  incidentType: string;
  containerWidth: number;
}

// safety incident detail markup
export default class SafetyIncident extends React.Component<ISafetyIncidentProps, localState> {

  // constructor
  public constructor(props: ISafetyIncidentProps) {
    super(props);
  }

  // data retrieval
  public componentDidMount() {
    const incidentId = this.retrieveIncidentId();
    const reactHandler = this;
    // these constants could be established in the app properties if desired
    const rootUrl = window.location.origin;
    const listName = this.props.listName;
    console.log("component did mount")

    jquery('.pageHeader').addClass('something');

    console.log("after first jquery call")
    
     //url: rootUrl + "/sites/apps/_api/web/lists/GetByTitle('" + listName + "')/Items(" + incidentId + ")",
    const url = rootUrl + "/sites/apps/_api/web/lists";
    console.log(url);
    jquery.ajax({
      url: url,
      type: "GET",
      dataType: "json",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: (resultData) => {
        console.log(resultData)
        /*
        reactHandler.setState({
          incidentId: resultData.d.Id,
          incidentDate: resultData.d.occurred,
          jobNumber: resultData.d.jobNumber,
          incidenype: resultData.d.incidentType,
          jobLocation: resultData.d.jobLocation,
          containerWidth: document.getElementsByClassName('CanvasSection')[0].clientWidth
        });
        */
      },
      error: (jqXHR, textStatus, errorThrown) => {
        console.log('jqXHR', jqXHR);
        console.log('text status', textStatus);
        console.log('error', errorThrown);
      }
    });
    /*
    // manual responsive adjustments
    jquery(window).resize(() => {
      let width = document.getElementsByClassName('CanvasSection')[0].clientWidth;
      reactHandler.setState({
        incidentId: reactHandler.state.incidentId,
        incidentDate: reactHandler.state.incidentDate,
        jobNumber: reactHandler.state.jobNumber,
        incidentType: reactHandler.state.incidentType,
        jobLocation: reactHandler.state.jobLocation,
        containerWidth: document.getElementsByClassName('CanvasSection')[0].clientWidth
      });
    });
    */
  }

  // extract the incident id from the url
  public retrieveIncidentId() {
    const pageName: string = location.href.split("/").slice(-1)[0];
    const startPos: number = pageName.indexOf('_') + 1;
    const endPos: number = pageName.indexOf('.aspx');
    const incidentId = pageName.substring(startPos, endPos);
    return incidentId;
  }

  // multi-line fields from SP lists come with html markup, this function removes some unnecessary tags
  public massageHTML(inputString) {
    let temp = inputString.replace(new RegExp('<br>', 'g'), '');
    let outputString = temp.replace(new RegExp('<p></p>', 'g'), '');
    return outputString;
  }

  public calculateColumnClasses(width) {
    if (width < 480) {
      return 'ms-Grid-col ms-sm12';
    } else if (width < 800) {
      return 'ms-Grid-col ms-sm6';
    } else {
      return 'ms-Grid-col ms-sm3';
    }
  }

  // render function
  public render(): React.ReactElement<ISafetyIncidentProps> {
    const thisIncident = this.state;

    if (!thisIncident) {
      return <div>Loading...</div>;
    }
    const columnClasses = this.calculateColumnClasses(this.state.containerWidth);
    return (
      <div className="header-container">
        <Fabric>
          <div>
            <div className="ms-Grid">
              <div className="ms-Grid-row">
                <div className={columnClasses}>
                  <div className={styles.incidentBox}>
                    <div className={styles.incidentBoxInner}>
                      <div className={styles.incidentLabel}>{thisIncident.jobNumber}</div>
                      <div className={styles.incidentLabelSm}>Job Number</div>
                    </div>
                  </div>
                </div>
                <div className={columnClasses}>
                  <div className={styles.incidentBox}>
                    <div className={styles.incidentBoxInner}>
                      <div className={styles.incidentLabel}>{thisIncident.incidentType}</div>
                      <div className={styles.incidentLabelSm}>Ocurred</div>
                    </div>
                  </div>
                </div>
                <div className={columnClasses}>
                  <div className={styles.incidentBox}>
                    <div className={styles.incidentBoxInner}>
                      <div className={styles.incidentLabel} dangerouslySetInnerHTML={{ __html: this.massageHTML(thisIncident.jobLocation) }}></div>
                      <div className={styles.incidentLabelSm}>Job Location</div>
                    </div>
                  </div>
                </div>
                <div className={columnClasses}>
                  <div className={styles.incidentBox}>
                    <div className={styles.incidentBoxInner}>
                      <div className={styles.incidentLabel}>{thisIncident.incidentType}</div>
                      <div className={styles.incidentLabelSm}>Type</div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </Fabric>
      </div >
    );
  }
}

