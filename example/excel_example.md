# Eexcel Example

API

- Address

  http://localhost:8080/excel

- Request

  ```
  {
    "fileName":"grocery_stock", 
    "sheet":[
        {
            "name":"fruit",
            "isRowTitle":false,
            "data":{
                "fruit":["apple","orange","strawberry"], 
                "stock":[2007,101,4899],
                "price":[1000,1500,500]
            }
        },
        {
            "name":"drink",
            "isRowTitle":true,
            "data":{
                "fruit":["water","juice","sprit"], 
                "stock":[2007,101,4899],
                "price":[1000,1500,500]
            }
        }
    ]
}

  ```

- result reference

  `grocery_stock.xlsx`
