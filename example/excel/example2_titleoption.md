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
            "isRowTitle": false,
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
                "backgroundColor": "AAAAFF"
            }
        }
    ]
  }
  ```

- result reference

  `example2_grocery_stock.xlsx`
