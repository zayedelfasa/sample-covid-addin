import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";
import Reducer from "../../reducer/reducer";
import { IS_INITIALIZED, IS_LOADING, IS_DONE, IS_ERROR } from '../../reducer/status';

// images references in the manifest
/* eslint-disable no-unused-vars */
import icon16 from "../../../assets/icon-16.png";
import icon32 from "../../../assets/icon-32.png";
import icon64 from "../../../assets/icon-64.png";
import icon80 from "../../../assets/icon-80.png";
import icon128 from "../../../assets/icon-128.png";
/* eslint-enable no-unused-vars */

/* global console, Excel, require */

const App = ({ isOfficeInitialized, title }) => {

  const [listItems, setListItems] = React.useState([]);
  const [stateListCovid, dispatch] = React.useReducer(Reducer, { type: IS_INITIALIZED, data: {}});
  const [stateListIDN, dispatchIDN] = React.useReducer(Reducer, { type: IS_INITIALIZED, data: {}});

  // get list
  React.useEffect(() => {
    setListItems([
      {
        icon: "Ribbon",
        primaryText: "Achieve more with Office integration",
      },
      {
        icon: "Unlock",
        primaryText: "Unlock features and functionality",
      },
      {
        icon: "Design",
        primaryText: "Create and visualize like a pro",
      },
    ]);
  }, []);

  const convertToTableExcel = async (dataWorld, dataIDN) => {
    try {
      await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();

        const header = [
          ["STATE", "CONFIRMED", "RECOVERED", "DEATHS"]
        ];

        let headerRange = sheet.getRange("B2:E4");
        headerRange.values = header;
        headerRange.format.fill.color = "#4472C4";
        headerRange.format.font.color = "white";

        let dataCovid = [
          ["World", `${dataWorld.confirmed.value}`, `${dataWorld.recovered.value}`, `${dataWorld.deaths.value}`],
          ["Indonesia", `${dataIDN.confirmed.value}`, `${dataIDN.recovered.value}`, `${dataIDN.deaths.value}`]
        ];

        let range = sheet.getRange("B3:D4");
        range.values = dataCovid;

        return context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  };

  const getDataCovid = async () => {
    dispatch({ type: IS_LOADING });
    dispatchIDN({ type: IS_LOADING });
    // fetch for the world
    await fetch("https://covid19.mathdro.id/api").then(async response => {
      dispatch({
        type: IS_DONE,
        data: response.json()
      });
    })
      .catch(error => {
        dispatch({
          type: IS_ERROR,
          error: error.toString()
        });
      });

    // fetch for IDN
    await fetch("https://covid19.mathdro.id/api/countries/IDN").then(async response => {
      dispatchIDN({
        type: IS_DONE,
        data: response.json()
      });
    })
      .catch(err => {
        dispatchIDN({
          type: IS_ERROR,
          error: err.toString()
        })
      });
  };

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    );
  }

  function checkIsDone() {
    return stateListCovid.type == IS_DONE && stateListIDN.type == IS_DONE;
  }

  function checkInitialized() {
    return stateListCovid.type == IS_INITIALIZED && stateListIDN.type == IS_INITIALIZED;
  }

  return (
    <div className="ms-welcome">
      <Header logo={require("./../../../assets/logo-filled.png")} title={title} message="Welcome" />
      <HeroList message="Discover what Office Add-ins can do for you today!" items={listItems}>
        <p className="ms-font-l">
          Modify the source files, then click <b>Run</b>.
        </p>
        {
          stateListCovid.loadState == IS_LOADING || stateListIDN.loadState == IS_LOADING ?
            <div>loading...</div>
            :
            <div>
              {
                <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={getDataCovid} >
                    Get data COVID                    
                </DefaultButton>
              }
              <div>
                {/* {stateListCovid.data == {} ? "" : stateListCovid.data} */}
                belum ada data
              </div>
            </div>
        }
      </HeroList>
    </div>
  );
}

export default App;
