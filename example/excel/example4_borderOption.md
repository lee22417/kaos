# Eexcel Example2 - titleOption

API

- Address

  http://localhost:8080/excel

- Request

  ```
    {
        "fileName": "grocery_stock",
        "sheet": [
            {
                "name": "fruit",
                "isRowTitle": true,
                "data": {
                    "fruit": [
                        "apple",
                        "orange",
                        "strawberry",
                        "grape",
                        "grapefruit",
                        "lemon"
                    ],
                    "stock": [
                        2007,
                        101,
                        4899,
                        45,
                        677,
                        810
                    ],
                    "price": [
                        1000,
                        1500,
                        500,
                        100,
                        500,
                        500
                    ]
                },
                "titleOption": {
                    "fontColor": "996633",
                    "fontSize": 20,
                    "bold": true,
                    "italic": true,
                    "backgroundColor": "EEDDFF"
                },
                "arrangeOption": {
                    "leftGap": 1,
                    "topGap": 3
                },
                "borderOption": {
                    "inside" : true,
                    "insideColor" : "coral",
                    "outside" : true,
                    "outsideColor" : "violet"
                }
            }
        ]
    }
  ```

- result reference

  `example4_grocery_stock.xlsx`
