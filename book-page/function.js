{
  "bindings": [
    {
      "authLevel": "Anonymous",
      "type": "httpTrigger",
      "direction": "in",
      "name": "req",
      "methods": ["get"],
      "route": "book/{slug}"
    },
    { "type": "http", "direction": "out", "name": "res" }
  ]
}
