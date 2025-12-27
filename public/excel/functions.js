{
  "functions"; [
    {
      "id": "MYADDIN",
      "name": "MYADDIN",
      "description": "Query AI about uploaded documents",
      "helpUrl": "https://word-addin-phi.vercel.app",
      "result": {
        "type": "string",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "query",
          "description": "Your question to the AI",
          "type": "string",
          "dimensionality": "scalar"
        }
      ],
      "options": {
        "stream": false,
        "cancelable": false
      }
    }
  ]
}