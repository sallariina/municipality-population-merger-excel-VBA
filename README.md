## Municipality population merger (Excel VBA)

VBA script that compiles population figures into a new Excel file.

* Statistics Finland's (Tilastokeskus) statistics on municipal population figures are used. Merges files by municipality.
* Saves new excel with timestamp to the script folder.

* Statistics downloaded from Tilastokeskus website: <br/>
Kuntien avainluvut / 2023 aluejaolla / Kuntien avainluvut 1987-2023 <br/>
https://pxdata.stat.fi/PxWeb/pxweb/fi/Kuntien_avainluvut/Kuntien_avainluvut__2023/kuntien_avainluvut_2023_aikasarja.px/

* Statistics used 2023 + 2022. The script currently only works for these years.

### How to use
1. Open Excel workbook
2. Open VBA editor. (Alt + F11)
3. Import the .bas file.
4. Add 2023 + 2022 files to the same folder as the Excel workbook
5. Run the macro
