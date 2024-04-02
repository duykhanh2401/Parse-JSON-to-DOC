const docx = require('docx');
const fs = require("fs")
const { Document, Packer, Paragraph, Table, TableCell, TableRow } = docx;

const data = {
    "code": 200,
    "message": "OK",
    "data": {
        "quote": null,
        "buyer": null,
        "seller": {
            "amount": "0.00",
            "shipping_fee": "0.0",
            "discount_amount": "0.00",
            "clone_price": "0.00",
            "tax_amount": "0.00",
            "payment_processing_fee": "0",
            "payment_info": null,
            "promotion_code": null
        },
        "fulfillment": "Unfulfilled",
        "state": "draft",
        "source": "custom.import",
        "currency": "USD",
        "note": null,
        "store_id": "WHBxEvQgzYAuUZUA",
        "channel": "fulfillment",
        "reference_id": "5-184",
        "shipping_method": "standard",
        "shipping": {
            "id": "Qbnq0vlApBYSPDVi",
            "email": "phongling77@gmail.com",
            "name": "Sal Carroll",
            "phone": null,
            "address": {
                "line1": "203 Huckleberry Ln",
                "line2": null,
                "city": "Duryea",
                "state": "PA",
                "postal_code": "18642",
                "country": "US",
                "country_name": null
            }
        },
        "extra_fee": "0",
        "items": [
            {
                "confirmed": true,
                "id": "nGrFdeTIBrewLVgp",
                "name": null,
                "product_id": "A27376-5e863e79828e44fa821ba74b1805a910",
                "product_type_id": null,
                "variant_id": null,
                "short_code": null,
                "color_value": null,
                "color_id": null,
                "color_name": null,
                "size_id": null,
                "size_name": null,
                "quantity": 1,
                "custom_data": null,
                "type": "custom",
                "unit_amount": null,
                "is_clone_design": false,
                "barcode": null,
                "notes": null,
                "designs": [
                    {
                        "id": "23217e10d878423582b99b177fe2c37d",
                        "type": "front",
                        "src": null,
                        "calculate_clone_price": true,
                        "resolution": null
                    }
                ],
                "mockups": [],
                "location": null,
                "sku": null,
                "ref_id": null,
                "position": 0,
                "sku_label": null,
                "mockup_api": null,
                "additional_designs": null,
                "amount": "0.00",
                "sub_amount": "0.00",
                "price": "0.00",
                "base_cost": "0.00",
                "clone_price": "0.00",
                "currency": "USD",
                "fulfillment_cost": null,
                "shipping_fee": "0.00",
                "tax_rate": "0.00",
                "tax_amount": "0.00",
                "payment_processing_fee": null,
                "shipping_method": "standard",
                "state": "approved",
                "shipping_method_allowed": null,
                "discount_amount": null,
                "tracking_codes": null,
                "trackings": [],
                "buyer_tax": null,
                "buyer_shipping_fee": null,
                "buyer_amount": null,
                "promotion_amount": null,
                "promotion_code": null,
                "promotion_code_amount": null,
                "promotion_auto_id": null,
                "promotion_auto_amount": null,
                "is_personalize": false,
                "quantity_buyer": 1
            }
        ],
        "payment_type": null,
        "ioss_number": null,
        "promotion_code": null,
        "already_applied_auto": false,
        "promotion_auto_id": null,
        "shipping_labels": [],
        "estimate_promotion": true,
        "shipping_method_buyer": null,
        "custom_label": null,
        "id": "A27376-CT-2158104",
        "store_name": null,
        "payment_state": "Incompleted",
        "fulfill_state": "Unfulfilled",
        "tracking_codes": null,
        "trackings": [],
        "callback_url": null,
        "domain": null,
        "create_date": "20231221T134150Z",
        "update_date": null,
        "order_date": null,
        "user_id": "A27376",
        "total_item": 1,
        "traffic_source": null,
        "shipping_method_allowed": null,
        "all_shipping_method": [],
        "shipping_methods": null,
        "require_refund": false,
        "promotion_amount": "0.00",
        "promotion_message": null,
        "is_promotion_code_valid": false,
        "fulfill_promotion_metadata": null,
        "confirmed_to_fulfill": true,
        "is_personalize": false
    }
}


numberCall = 0

function getType(value) {
    if (Array.isArray(value)) {
        return "array"
    }

    if (value == null) {
        return "null"
    }

    return typeof value
}

function detectData(data, subIndex) {
    numberCall++
    // console.log("Number Call: ", numberCall, data, typeof data);

    const arrayData = []
    let index = 0
    if (typeof data == 'object' && data) {
        console.log("------------------------");

        for ([key, value] of Object.entries(data)) {
            index++
            stt = subIndex ? subIndex + "." + index : String(index)
            arrayData.push({
                stt: stt,
                name: key,
                type: getType(value)
            })
            console.log(key, stt, index, typeof value);

            if (Array.isArray(value) && value.length) {
                arrayData.push(...detectData(value[0], stt))
            } else if (typeof value == 'object') {
                arrayData.push(...detectData(value, stt))
            }
        }
    }

    return arrayData
}

const res = detectData(data)

const rows = []

for (ele of res) {
    console.log(ele);
    rows.push(new TableRow({
        children: [
            new TableCell({ children: [new Paragraph(ele.stt)] }),
            new TableCell({ children: [new Paragraph(ele.name)] }),
            new TableCell({ children: [new Paragraph(ele.type)] }),
            new TableCell({ children: [] })
        ]
    }))
}

const table = new Table({
    rows
})

console.log(rows);


const doc = new Document({
    sections: [{
        children: [table],
    }],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});