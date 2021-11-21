import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";
import Reducer from "../../reducer/reducer";
import { IS_INITIALIZED, IS_LOADING, IS_DONE, IS_ERROR } from '../../reducer/status';

const App = ({ isOfficeInitialized, title }) => {

  const [listItems, setListItems] = React.useState([]);
  const [stateListCovid, dispatch] = React.useReducer(Reducer, { type: IS_INITIALIZED, data: {} });
  const [stateListIDN, dispatchIDN] = React.useReducer(Reducer, { type: IS_INITIALIZED, data: {} });

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

  React.useEffect(() => {
    if (stateListCovid.loadState == IS_DONE && stateListIDN.loadState == IS_DONE ) 
      convertToTableExcel();
  }, [stateListCovid.loadState, stateListIDN.loadState]);

  const convertToTableExcel = async () => {
    try {
      let dataWorld, dataIDN;
      dataWorld = stateListCovid.data;
      dataIDN = stateListIDN.data;

      console.log(`data world ${dataWorld.confirmed.value}`);

      await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();

        const header = [
          ["STATE", "CONFIRMED", "RECOVERED", "DEATHS"]
        ];

        let headerRange = sheet.getRange("B2:E2");
        headerRange.values = header;
        headerRange.format.fill.color = "#4472C4";
        headerRange.format.font.color = "white";

        let dataCovid = [
          ["World", `${dataWorld.confirmed.value}`, `${dataWorld.recovered.value}`, `${dataWorld.deaths.value}`],
          ["Indonesia", `${dataIDN.confirmed.value}`, `${dataIDN.recovered.value}`, `${dataIDN.deaths.value}`]
        ];

        let range = sheet.getRange("B3:E4");
        range.numberFormat=[["0"]]
        range.format.autofitRows = true;
        range.format.autofitColumns = true;
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

    await fetch("https://covid19.mathdro.id/api")
      .then(async (res) => {
        let data = await res.json();
        console.log(`Data Covid : ${data.confirmed.value}`);
        dispatch({
          type: IS_DONE,
          data: data
        });
      })
      .catch(err => {
        dispatch({
          type: IS_ERROR,
          error: err
        });
      });

    // fetch for IDN
    await fetch("https://covid19.mathdro.id/api/countries/IDN")
      .then(async (res) => {
        let data = await res.json();
        dispatchIDN({
          type: IS_DONE,
          data: data
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
            </div>
        }
      </HeroList>
    </div>
  );
}

export default App;
