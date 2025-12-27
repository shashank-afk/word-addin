console.log("functions.js loaded");

/**
 * @customfunction
 */
async function ASKADDIN(question) {
  console.log("ASKADDIN called", question);
  return "Hello " + question;
}
