# Excel File Summary

## Summary

- **Sheet1**: 2 rows, 2 columns
- **Sheet2**: 2 rows, 2 columns
- **Sheet3**: 2 rows, 2 columns
- **Sheet4**: 2 rows, 2 columns
- **Sheet5**: 2 rows, 2 columns
- **Sheet6**: 2 rows, 2 columns


## Sheet: Sheet1

Dimensions: 2 rows × 2 columns

| Key | Value |
| --- | --- |
| A  | Apple |


## Sheet: Sheet2

Dimensions: 2 rows × 2 columns

| Key | Value |
| --- | --- |
| C | Banana |


## Sheet: Sheet3

Dimensions: 2 rows × 2 columns

| Key  | Value |
| --- | --- |
| E | Orange |


## Sheet: Sheet4

Dimensions: 2 rows × 2 columns

| Key  | Value |
| --- | --- |
| B | Mango |


## Sheet: Sheet5

Dimensions: 2 rows × 2 columns

| Key | Value |
| --- | --- |
| D | Watermelon |


## Sheet: Sheet6

Dimensions: 2 rows × 2 columns

| Key | Value |
| --- | --- |
| H | #VALUE! |


## Images

### Image Analysis

Found 1 embedded image in the Excel file:

**image1.png**
- **Location**: Embedded as rich value (not tied to specific sheet/cell via traditional drawing objects)  
- **Dimensions**: 360 × 360 pixels
- **Format**: PNG (palette mode with transparency)
- **Description**: The image appears to be a graphic with a dark black border and light brown/beige center area (RGB approximately 195, 189, 177). Based on the analysis of Excel's internal structure, this image is embedded as a "_localImage" rich value type rather than a traditional drawing object.
- **File**: ![image1.png](media/image1.png)

### Technical Notes

The image was successfully extracted using the robust unzip approach from `xl/media/image1.png`. It's referenced through Excel's rich data system (`xl/richData/`) but not mapped to specific cells through traditional drawing relationships. This type of embedding is why the image couldn't be detected through openpyxl's standard `_images` attribute.