const XLSX = require("xlsx");
const axios = require("axios");
const Promise = require("bluebird");

fs = require("fs");
const workbook = XLSX.readFile("./in/someSpreadsheet.xlsx");
const sheet_name_list = workbook.SheetNames;

const xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

const rootUrl = "https://www.w3.org";

let errors = 0;
let success = 0;
let count = 1;
let failureArray = [
  {
    ROW: "",
    FROM: "",
    TO: "",
    CODE: "",
    NUMBER_OF_REDIRECTS: "",
    ERROR_MESSAGE: "",
    REQ_URL: "",
  },
];
console.time("Total Time");

Promise.map(
  xlData,

  async (col, idx) => {
    console.log(`Process: ${count}, Row ${idx + 2}`);
    count++;
    let obj;

    const from = col["source"];
    const to = col["target"];

    let failure = false;
    let responseCode = "";
    let numberOfRedirects;

    obj = await axios(encodeURI(`${rootUrl}${from}`))
      .then((response) => {
        success++;

        responseCode = response.status ? response.status : "";

        failure = false;

        numberOfRedirects =
          response && response.request && response.request._redirectable
            ? response.request._redirectable._redirectCount
            : "";

        return {
          ...col,
          ["Failure"]: failure,
          ["Response Code"]: responseCode,
          ["Number of Redirects"]: numberOfRedirects,
        };
      })
      .catch((err) => {
        console.log(`ROW: ${idx + 2}, ERROR:${err.message}`);
        errors++;
        const response = err.response ? err.response : "";

        failure = true;

        //   if (idx === whatever row you want to validate) {
        //     fs.writeFile(
        //       "testJson.json",
        //       JSON.safeStringify(response),
        //       function (err) {
        //         if (err) return console.log(err);
        //       }
        //     );
        //   }

        numberOfRedirects =
          response && response.request && response.request._redirectable
            ? response.request._redirectable._redirectCount
            : "";

        failureArray.push({
          ROW: idx + 2,
          FROM: from,
          TO: to,
          CODE: response && response.status ? response.status : "",
          NUMBER_OF_REDIRECTS: numberOfRedirects,
          ERROR_MESSAGE: err.message ? err.message : "",
          REQ_URL:
            response.request && response.request.path
              ? response.request.path
              : "",
        });

        return {
          ...col,
          ["Failure"]: failure,
          ["Response Code"]: response && response.status ? response.status : "",
          ["Number of Redirects"]: numberOfRedirects,
        };
      });

    return obj;
  },
  { concurrency: 10 }
).then((data) => {
  const output = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, output, "All Redirects");
  XLSX.writeFile(wb, "./out/allRedirects.xlsx");

  const failureOutput = XLSX.utils.json_to_sheet(failureArray);
  const failBook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(failBook, failureOutput, "Failed Redirects");
  XLSX.writeFile(failBook, "./out/failedRedirects.xlsx");

  console.log(`Successful redirects: ${success}, Unsuccesful: ${errors} `);
  console.timeEnd("Total Time");
});

// safely handles circular references
JSON.safeStringify = (obj, indent = 2) => {
  let cache = [];
  const retVal = JSON.stringify(
    obj,
    (key, value) =>
      typeof value === "object" && value !== null
        ? cache.includes(value)
          ? undefined // Duplicate reference found, discard key
          : cache.push(value) && value // Store value in our collection
        : value,
    indent
  );
  cache = null;
  return retVal;
};
