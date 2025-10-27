{
  "bindings": [
    {
      "authLevel": "Anonymous",
      "type": "httpTrigger",
      "direction": "in",
      "name": "req",
      "methods": ["get"],
      "route": "book/{slug}/page"
    },
    { "type": "http", "direction": "out", "name": "res" }
  ]
}
