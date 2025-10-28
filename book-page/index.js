module.exports = async function (context, req) {
  context.log("Handler executing");
  context.res = {
    status: 200,
    headers: { 'Content-Type': 'text/plain' },
    body: 'Hello World'
  };
};
