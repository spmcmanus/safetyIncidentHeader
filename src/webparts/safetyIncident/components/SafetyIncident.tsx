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
    const rootUrl = window.location.origin;
    const listName = this.props.listName;
    const siteName = this.props.siteName;
    // build out url
    const url = rootUrl + "/sites/" + siteName + "/_api/web/lists/GetByTitle('" + listName + "')/Items(" + incidentId + ")";
    jquery.ajax({
      url: url,
      type: "GET",
      dataType: "json",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: (resultData) => {
        reactHandler.setState({
          incidentId: resultData.d.Id,
          incidentDate: resultData.d.Occurred1, //occurred,
          jobNumber: resultData.d.Job_x0020_Number,  //jobNumber,
          incidentType: resultData.d.Incident_x0020_Type, //incidentType,
          jobLocation: resultData.d.Title,//jobLocation,
          containerWidth: 1000
        });
      },
      error: (jqXHR, textStatus, errorThrown) => {
        console.log('jqXHR', jqXHR);
        console.log('text status', textStatus);
        console.log('error', errorThrown);
      },
      complete: () => {
        window.dispatchEvent(new Event('resize'));
      }
    });
    
    // manual responsive adjustments
    jquery(window).resize(() => {
      const section = jquery('.header-container');
      let width = 1000;
      if (section.length > 0 ) {
        width = section.closest('.CanvasSection')[0].clientWidth;
      } 
      reactHandler.setState({
        incidentId: reactHandler.state.incidentId,
        incidentDate: reactHandler.state.incidentDate,
        jobNumber: reactHandler.state.jobNumber,
        incidentType: reactHandler.state.incidentType,
        jobLocation: reactHandler.state.jobLocation,
        containerWidth: width
      });
    });
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

