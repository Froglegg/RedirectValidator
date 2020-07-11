const XLSX = require("xlsx");
const axios = require("axios");
const Promise = require("bluebird");

fs = require("fs");
const workbook = XLSX.readFile("./in/careRedirects.xlsx");
const sheet_name_list = workbook.SheetNames;

const xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

const rootUrl = "";

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
    BAD_URL: "",
    ERROR_MESSAGE: "",
  },
];
console.time("Total Time");
Promise.map(
  xlData,
  async (col, idx) => {
    console.log(`Process: ${count}, Row ${idx + 1}`);
    count++;
    let obj;

    const from = col["FROM"];
    const to = col["TO"];

    if (idx === 0) {
      obj = {
        ...col,
        ["__EMPTY_7"]: "Failure",
        ["__EMPTY_8"]: "Response Code",
        ["__EMPTY_9"]: "Number of Redirects",
        ["__EMPTY_10"]: "Redirect URL != Target",
      };
    } else {
      let failure = false;
      let responseCode = "";
      let numberOfRedirects;
      let wrongRedirectUrl = false;

      obj = await axios(encodeURI(`${rootUrl}${from}`))
        .then((response) => {
          success++;
          if (
            response.request.path !== to &&
            response.request.path !== `${to}/`
          ) {
            wrongRedirectUrl = true;
          }

          responseCode = response.status;

          failure = response.status !== 200 ? true : false;
          numberOfRedirects = response.request._redirectable._redirectCount;

          return {
            ...col,
            ["__EMPTY_7"]: failure,
            ["__EMPTY_8"]: responseCode,
            ["__EMPTY_9"]: numberOfRedirects,
            ["__EMPTY_10"]: wrongRedirectUrl,
          };
        })
        .catch((err) => {
          console.log(`ROW: ${idx + 2}, ERROR:${err.message}`);
          errors++;
          const response = err.response ? err.response : "";

          if (
            response &&
            response.request.path !== to &&
            response.request.path !== `${to}/`
          ) {
            wrongRedirectUrl = true;
          }

          responseCode = response.status ? response.status : "N/A";

          failure = response.status !== 200 ? true : false;
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
              : 0;

          failureArray.push({
            ROW: idx + 2,
            FROM: from,
            TO: to,
            CODE: response.status,
            NUMBER_OF_REDIRECTS: numberOfRedirects,
            BAD_URL: wrongRedirectUrl,
            ERROR_MESSAGE: err.message ? err.message : "",
          });

          return {
            ...col,
            ["__EMPTY_7"]: failure,
            ["__EMPTY_8"]: responseCode,
            ["__EMPTY_9"]: numberOfRedirects,
            ["__EMPTY_10"]: wrongRedirectUrl,
          };
        });
    }
    return obj;
  },
  { concurrency: 10 }
).then((data) => {
  const output = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, output, "testSheet");
  XLSX.writeFile(wb, "./out/example.xlsx");

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
