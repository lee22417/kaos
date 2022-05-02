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
                        "strawberry"
                    ],
                    "stock": [
                        2007,
                        101,
                        4899
                    ],
                    "price": [
                        1000,
                        1500,
                        500
                    ]
                },
                "titleOption": {
                    "fontColor": "993366",
                    "fontSize": 20,
                    "bold": true,
                    "italic": true,
                    "backgroundColor": "AAAAFF"
                },
                "arrangeOption": {
                    "leftGap": 1,
                    "topGap": 3
                }
            }
        ]
    }
  ```

- result reference

  `example3_grocery_stock.xlsx`
