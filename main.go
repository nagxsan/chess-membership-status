package main

import (
	"compress/gzip"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"
	"strconv"
	"strings"
	"time"

	excelize "github.com/xuri/excelize/v2"
)

func getMembership(id string) (bool, error) {
	url := fmt.Sprintf("https://admin.aicf.in/api/players?name=%s&state=0&city=0", id)

	response, err := http.Get(url)
	if err != nil {
		return false, fmt.Errorf("Error in making the API call: %v", err)
	}

	defer response.Body.Close()

	var data map[string]interface{}
	if err := json.NewDecoder(response.Body).Decode(&data); err != nil {
		return false, fmt.Errorf("Error in decoding response body: %v", err)
	}

	dataArray, ok := data["data"].([]interface{})
	if !ok || len(dataArray) <= 0 {
		return false, fmt.Errorf("Error in getting dataArray: %v", err)
	}

	if len(dataArray) != 1 {
		return false, fmt.Errorf("Incorrect ID")
	}

	playerData, ok := dataArray[0].(map[string]interface{})
	if !ok {
		return false, fmt.Errorf("Error in getting playerData: %v", err)
	}

	membershipStatus, ok := playerData["membership_status"].(bool)
	if !ok {
		return false, fmt.Errorf("Error in getting membershipStatus: %v", err)
	}

	return membershipStatus, nil
}

func getMCAMembershipStatus(id string) (bool, error) {
	headers := map[string]string{
		"Cache-Control":   "no-cache",
		"User-Agent":      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36",
		"Accept":          "*/*",
		"Accept-Encoding": "gzip, deflate, br",
	}

	url := fmt.Sprintf("https://mcachess.in/Tournament_Registration/fetch_registrarion_type_web.php?page=1&query=%s", id)

	client := &http.Client{}
	req, err := http.NewRequest("GET", url, nil)
	if err != nil {
		return false, fmt.Errorf("Error in creating new request: %v", err)
	}

	for key, value := range headers {
		req.Header.Set(key, value)
	}

	response, err := client.Do(req)
	if err != nil {
		return false, fmt.Errorf("Error in making an API call: %v", err)
	}
	defer response.Body.Close()

	if response.StatusCode != http.StatusOK {
		return false, fmt.Errorf("Unexpected status code; API call failed")
	}

	var reader io.ReadCloser
	switch response.Header.Get("Content-Encoding") {
		case "gzip":
			reader, err = gzip.NewReader(response.Body)
			if err != nil {
				return false, fmt.Errorf("Error in unzipping response body: %v", err)
			}
			defer reader.Close()
		default:
			reader = response.Body
	}

	body, err := io.ReadAll(reader)
	if err != nil {
		return false, fmt.Errorf("Error in reading response body: %v", err)
	}

	responseHTML := string(body)
	findPos := strings.Index(responseHTML, "</label>")
	if findPos == -1 {
		return false, nil
	}

	getResponseNum, err := strconv.Atoi(responseHTML[findPos-1:findPos])
	if err != nil {
		return false, fmt.Errorf("Error in getting the MCA status response: %v", err)
	}

	if getResponseNum == 1 {
		return true, nil
	} else {
		return false, nil
	}
}

func main() {
	startTime := time.Now()

	var excelPath, sheetName string
	var tableRowNumber int
	fmt.Print("Enter excel file path with extension (no spaces anywhere in the file path): ")
	_, err := fmt.Scan(&excelPath)
	if err != nil {
		fmt.Println("Error in excel file name: ", err)
		os.Exit(1)
	}

	fmt.Print("Enter sheet name: ")
	_, err = fmt.Scan(&sheetName)
	if err != nil {
		fmt.Println("Error in sheet name: ", err)
		os.Exit(1)
	}

	fmt.Print("Enter row number where table starts: ")
	_, err = fmt.Scan(&tableRowNumber)
	if err != nil {
		fmt.Println("Error in table row number: ", err)
		os.Exit(1)
	}

	xl, err := excelize.OpenFile(excelPath)
	if err != nil {
		fmt.Println("Error opening file: ", err)
		os.Exit(1)
	}

	defer xl.Close()

	sheetDimension, err := xl.GetSheetDimension(sheetName)
	if err != nil {
		fmt.Println("Error getting sheet dimension: ", err)
		os.Exit(1)
	}

	lastCol := strings.Split(sheetDimension, ":")[1][0]

	membershipStatusCell := string(lastCol+1) + strconv.Itoa(tableRowNumber)
	membershipStatusMCACell := string(lastCol+2) + strconv.Itoa(tableRowNumber)

	err = xl.SetCellStr(sheetName, membershipStatusCell, "Membership_Status")
	if err != nil {
		fmt.Println("Error setting the Membership_Status column: ", err)
		os.Exit(1)
	}

	err = xl.SetCellStr(sheetName, membershipStatusMCACell, "Membership_Status_MCA")
	if err != nil {
		fmt.Println("Error setting the Membership_Status_MCA column: ", err)
		os.Exit(1)
	}

	err = xl.Save()
	if err != nil {
		fmt.Println("Error saving excel file: ", err)
		os.Exit(1)
	}

	membershipStatusColumn := string(lastCol + 1)
	membershipStatusMCAColumn := string(lastCol + 2)

	var aicfColumn, fideColumn, mcaColumn string

	for colNum := byte(0); (colNum + 65) <= lastCol; colNum++ {
		colName := string(colNum + 65)
		cellNum := colName + strconv.Itoa(tableRowNumber)
		cellValue, err := xl.GetCellValue(sheetName, cellNum)
		if err != nil {
			fmt.Println("Error reading coliumn cell value: ", err)
			os.Exit(1)
		}

		if cellValue == "AICF ID" {
			aicfColumn = colName
		} else if cellValue == "FIDE ID" {
			fideColumn = colName
		} else if cellValue == "MCA ID" {
			mcaColumn = colName
		}
	}

	for dataRowNumber := tableRowNumber + 1; ; dataRowNumber++ {
		firstCellValue, err := xl.GetCellValue(sheetName, "A" + strconv.Itoa(dataRowNumber))
		if err != nil {
			fmt.Println("Error fetching first cell value to verify the table end: ", err)
			os.Exit(1)
		}

		if firstCellValue == "" {
			break
		}

		for colNum := byte(0); (colNum + 65) <= lastCol; colNum++ {
			colName := string(colNum + 65)
			if colName == aicfColumn || colName == fideColumn {

				membershipStatusCellValue, err := xl.GetCellValue(sheetName, membershipStatusColumn + strconv.Itoa(dataRowNumber))
				if err != nil {
					fmt.Println("Error getting membership status cell value: ", err)
					os.Exit(1)
				}

				if membershipStatusCellValue == "Yes" {
					continue
				}

				id, err := xl.GetCellValue(sheetName, colName + strconv.Itoa(dataRowNumber))
				if err != nil {
					fmt.Println("Error getting cell value: ", err)
					os.Exit(1)
				}

				membershipStatus, err := getMembership(id)
				if err != nil {
					fmt.Printf("ID: %v; Error in getting membership status response: %v\n", id, err)
				}

				time.Sleep(2 * time.Second)

				if membershipStatus == false || err != nil {
					err = xl.SetCellStr(sheetName, membershipStatusColumn + strconv.Itoa(dataRowNumber), "Check Manually")
					if err != nil {
						fmt.Println("Error setting membership status value: ", err)
						os.Exit(1)
					}
				} else {
					err = xl.SetCellStr(sheetName, membershipStatusColumn + strconv.Itoa(dataRowNumber), "Yes")
					if err != nil {
						fmt.Println("Error setting membership status value: ", err)
						os.Exit(1)
					}
				}
			} else if colName == mcaColumn {
				id, err := xl.GetCellValue(sheetName, colName + strconv.Itoa(dataRowNumber))
				if err != nil {
					fmt.Println("Error getting cell value: ", err)
					os.Exit(1)
				}

				if strings.HasPrefix(id, "PO") == false {
					err = xl.SetCellStr(sheetName, membershipStatusMCAColumn + strconv.Itoa(dataRowNumber), "Check Manually")
					if err != nil {
						fmt.Println("Error setting membership status MCA value: ", err)
						os.Exit(1)
					}
					continue
				}

				membershipStatusMCA, err := getMCAMembershipStatus(id)
				if err != nil {
					fmt.Printf("ID: %v ; Error in getting MCA membership status response: %v\n", id, err)
				}

				time.Sleep(2 * time.Second)

				if membershipStatusMCA == false || err != nil {
					err = xl.SetCellStr(sheetName, membershipStatusMCAColumn + strconv.Itoa(dataRowNumber), "Check Manually")
					if err != nil {
						fmt.Println("Error setting membership status MCA value: ", err)
						os.Exit(1)
					}
				} else {
					err = xl.SetCellStr(sheetName, membershipStatusMCAColumn + strconv.Itoa(dataRowNumber), "Yes")
					if err != nil {
						fmt.Println("Error setting membership status MCA value: ", err)
						os.Exit(1)
					}
				}
			}
		}
	}

	err = xl.Save()
	if err != nil {
		fmt.Println("Error saving excel file: ", err)
		os.Exit(1)
	}
	
	endTime := time.Now()

	fmt.Println("Time taken: ", endTime.Sub(startTime))

	os.Exit(0)
}
