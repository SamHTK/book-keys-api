const { TableClient } = require("@azure/data-tables");

function table(tableName) {
  const conn = process.env.AzureWebJobsStorage;
  if (!conn) throw new Error("AzureWebJobsStorage not configured");
  return TableClient.fromConnectionString(conn, tableName);
}

async function ensureTables() {
  const t1 = table("requests");
  const t2 = table("requestsByUser");
  await t1.createTable().catch(() => {});
  await t2.createTable().catch(() => {});
}

async function putRequestEntity(entity) {
  const t = table("requests");
  await t.upsertEntity(entity, "Replace");
}

async function getRequestEntity(requestId) {
  const t = table("requests");
  // PartitionKey = requestId, RowKey = "r"
  return await t.getEntity(requestId, "r");
}

async function putRequestByUser(ownerUserId, requestId, summary) {
  const t = table("requestsByUser");
  await t.upsertEntity({
    partitionKey: ownerUserId,
    rowKey: requestId,
    ...summary
  }, "Replace");
}

async function updateRequestStatus(requestId, patch) {
  const t = table("requests");
  const ent = await t.getEntity(requestId, "r");
  Object.assign(ent, patch);
  await t.updateEntity(ent, "Replace");
}

module.exports = {
  ensureTables,
  putRequestEntity,
  getRequestEntity,
  putRequestByUser,
  updateRequestStatus
};
