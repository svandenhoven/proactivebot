import app from "./app";

// Start the application
(async () => {
  await app.start();
  //await app.send("2f575031-4e08-4770-bc63-bfe48a4a62ad", "Hello World!");
  console.log(`\nBot started, app listening to`, process.env.PORT || process.env.port || 3978);
})();
