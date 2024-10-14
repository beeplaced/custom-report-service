/** Constructs a drawing object for an image to be used in a DOCX document.
 * This function creates a structured object representing the drawing of an image,
 * based on the provided parameters. The resulting object conforms to the expected
 * format for embedding images in DOCX documents.
 * @param {Object} entry - The entry object containing parameters for the drawing object.
 * @param {string} entry.imageID - The unique identifier for the image.
 * @param {string} entry.rID - The relationship ID for the image.
 * @param {string} entry.cellWidth - The width of the cell where the image will be placed (in EMU).
 * @param {string} entry.cellHeight - The height of the cell where the image will be placed (in EMU).
 * @returns {Object} The constructed drawing object for the image.
 */
module.exports.buildDrawingObject = (entry) => {
    const { imageID, rID, cellWidth, cellHeight } = entry

    return {
        "type": "element",
        "name": "w:p",
        "attributes": {
            "w:rsidR": "00000000",
            "w:rsidDel": "00000000",
            "w:rsidP": "00000000",
            "w:rsidRDefault": "00000000",
            "w:rsidRPr": "00000000",
            "w14:paraId": "00000002"
        },
        "elements": [
            {
                "type": "element",
                "name": "w:pPr",
                "elements": [
                    {
                        "type": "element",
                        "name": "w:keepNext",
                        "attributes": {
                            "w:val": "0"
                        }
                    },
                    {
                        "type": "element",
                        "name": "w:keepLines",
                        "attributes": {
                            "w:val": "0"
                        }
                    },
                    {
                        "type": "element",
                        "name": "w:widowControl",
                        "attributes": {
                            "w:val": "0"
                        }
                    },
                    {
                        "type": "element",
                        "name": "w:pBdr",
                        "elements": [
                            {
                                "type": "element",
                                "name": "w:top",
                                "attributes": {
                                    "w:space": "0",
                                    "w:sz": "0",
                                    "w:val": "nil"
                                }
                            },
                            {
                                "type": "element",
                                "name": "w:left",
                                "attributes": {
                                    "w:space": "0",
                                    "w:sz": "0",
                                    "w:val": "nil"
                                }
                            },
                            {
                                "type": "element",
                                "name": "w:bottom",
                                "attributes": {
                                    "w:space": "0",
                                    "w:sz": "0",
                                    "w:val": "nil"
                                }
                            },
                            {
                                "type": "element",
                                "name": "w:right",
                                "attributes": {
                                    "w:space": "0",
                                    "w:sz": "0",
                                    "w:val": "nil"
                                }
                            },
                            {
                                "type": "element",
                                "name": "w:between",
                                "attributes": {
                                    "w:space": "0",
                                    "w:sz": "0",
                                    "w:val": "nil"
                                }
                            }
                        ]
                    },
                    {
                        "type": "element",
                        "name": "w:shd",
                        "attributes": {
                            "w:fill": "auto",
                            "w:val": "clear"
                        }
                    },
                    {
                        "type": "element",
                        "name": "w:spacing",
                        "attributes": {
                            "w:after": "0",
                            "w:before": "0",
                            "w:line": "240",
                            "w:lineRule": "auto"
                        }
                    },
                    {
                        "type": "element",
                        "name": "w:ind",
                        "attributes": {
                            "w:left": "0",
                            "w:right": "0",
                            "w:firstLine": "0"
                        }
                    },
                    {
                        "type": "element",
                        "name": "w:jc",
                        "attributes": {
                            "w:val": "left"
                        }
                    },
                    {
                        "type": "element",
                        "name": "w:rPr"
                    }
                ]
            },
            {
                "type": "element",
                "name": "w:r",
                "attributes": {
                    "w:rsidDel": "00000000",
                    "w:rsidR": "00000000",
                    "w:rsidRPr": "00000000"
                },
                "elements": [
                    {
                        "type": "element",
                        "name": "w:rPr"
                    },
                    {
                        "type": "element",
                        "name": "w:drawing",
                        "elements": [
                            {
                                "type": "element",
                                "name": "wp:inline",
                                "attributes": {
                                    "distB": "114300",
                                    "distT": "114300",
                                    "distL": "114300",
                                    "distR": "114300"
                                },
                                "elements": [
                                    {
                                        "type": "element",
                                        "name": "wp:extent",
                                        "attributes": {
                                            "cx": cellWidth,
                                            "cy": cellHeight
                                        }
                                    },
                                    {
                                        "type": "element",
                                        "name": "wp:effectExtent",
                                        "attributes": {
                                            "b": "0",
                                            "l": "0",
                                            "r": "0",
                                            "t": "0"
                                        }
                                    },
                                    {
                                        "type": "element",
                                        "name": "wp:docPr",
                                        "attributes": {
                                            "id": `${imageID}`,
                                            "name": `image${imageID}.png`
                                        }
                                    },
                                    {
                                        "type": "element",
                                        "name": "a:graphic",
                                        "elements": [
                                            {
                                                "type": "element",
                                                "name": "a:graphicData",
                                                "attributes": {
                                                    "uri": "http://schemas.openxmlformats.org/drawingml/2006/picture"
                                                },
                                                "elements": [
                                                    {
                                                        "type": "element",
                                                        "name": "pic:pic",
                                                        "elements": [
                                                            {
                                                                "type": "element",
                                                                "name": "pic:nvPicPr",
                                                                "elements": [
                                                                    {
                                                                        "type": "element",
                                                                        "name": "pic:cNvPr",
                                                                        "attributes": {
                                                                            "id": `${imageID}`,
                                                                            "name": `image${imageID}.png`
                                                                        }
                                                                    },
                                                                    {
                                                                        "type": "element",
                                                                        "name": "pic:cNvPicPr",
                                                                        "attributes": {
                                                                            "preferRelativeResize": "0"
                                                                        }
                                                                    }
                                                                ]
                                                            },
                                                            {
                                                                "type": "element",
                                                                "name": "pic:blipFill",
                                                                "elements": [
                                                                    {
                                                                        "type": "element",
                                                                        "name": "a:blip",
                                                                        "attributes": {
                                                                            "r:embed": `${rID}`
                                                                        }
                                                                    },
                                                                    {
                                                                        "type": "element",
                                                                        "name": "a:srcRect",
                                                                        "attributes": {
                                                                            "b": "0",
                                                                            "l": "0",
                                                                            "r": "0",
                                                                            "t": "0"
                                                                        }
                                                                    },
                                                                    // {
                                                                    //     "type": "element",
                                                                    //     "name": "a:stretch",
                                                                    //     "elements": [
                                                                    //         {
                                                                    //             "type": "element",
                                                                    //             "name": "a:fillRect"
                                                                    //         }
                                                                    //     ]
                                                                    // }
                                                                ]
                                                            },
                                                            {
                                                                "type": "element",
                                                                "name": "pic:spPr",
                                                                "elements": [
                                                                    {
                                                                        "type": "element",
                                                                        "name": "a:xfrm",
                                                                        "elements": [
                                                                            {
                                                                                "type": "element",
                                                                                "name": "a:off",
                                                                                "attributes": {
                                                                                    "x": "0",
                                                                                    "y": "0"
                                                                                }
                                                                            },
                                                                            {
                                                                                "type": "element",
                                                                                "name": "a:ext",
                                                                                "attributes": {
                                                                                    "cx": cellWidth,
                                                                                    "cy": cellHeight
                                                                                }
                                                                            }
                                                                        ]
                                                                    },
                                                                    {
                                                                        "type": "element",
                                                                        "name": "a:prstGeom",
                                                                        "attributes": {
                                                                            "prst": "rect"
                                                                        }
                                                                    },
                                                                    {
                                                                        "type": "element",
                                                                        "name": "a:ln"
                                                                    }
                                                                ]
                                                            }
                                                        ]
                                                    }
                                                ]
                                            }
                                        ]
                                    }
                                ]
                            }
                        ]
                    }
                ]
            },
            // {
            //     "type": "element",
            //     "name": "w:r",
            //     "attributes": {
            //         "w:rsidDel": "00000000",
            //         "w:rsidR": "00000000",
            //         "w:rsidRPr": "00000000"
            //     },
            //     "elements": [
            //         {
            //             "type": "element",
            //             "name": "w:rPr",
            //             "elements": [
            //                 {
            //                     "type": "element",
            //                     "name": "w:rtl",
            //                     "attributes": {
            //                         "w:val": "0"
            //                     }
            //                 }
            //             ]
            //         }
            //     ]
            // }
        ]
    }
}