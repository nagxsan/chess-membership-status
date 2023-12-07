package main

import (
	"compress/gzip"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"
	"regexp"
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

func getMCAId(id string) (string, error) {
	if id == "" {
		return "", fmt.Errorf("Error: blank AICF ID")
	}
	
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
		return "", fmt.Errorf("Error in creating new request: %v", err)
	}

	for key, value := range headers {
		req.Header.Set(key, value)
	}

	response, err := client.Do(req)
	if err != nil {
		return "", fmt.Errorf("Error in making an API call: %v", err)
	}
	defer response.Body.Close()

	if response.StatusCode != http.StatusOK {
		return "", fmt.Errorf("Unexpected status code; API call failed")
	}

	var reader io.ReadCloser
	switch response.Header.Get("Content-Encoding") {
		case "gzip":
			reader, err = gzip.NewReader(response.Body)
			if err != nil {
				return "", fmt.Errorf("Error in unzipping response body: %v", err)
			}
			defer reader.Close()
		default:
			reader = response.Body
	}

	body, err := io.ReadAll(reader)
	if err != nil {
		return "", fmt.Errorf("Error in reading response body: %v", err)
	}

	responseHTML := string(body)
	re := regexp.MustCompile(`PO.*<a>`)
	matches := re.FindAllString(responseHTML, -1)

	if len(matches) > 0 {
		var mcaIds string
		for i, mcaId := range matches {
			matches[i] = mcaId[0:len(mcaId) - 3]
		}
		mcaIds = strings.Join(matches, ", ")
		return mcaIds, nil
	} else {
		return "", fmt.Errorf("Error: no MCA IDs returned for the given AICF ID: %v", err)
	}
}

func getAICFId(id string) (string, error) {
	url := fmt.Sprintf("https://admin.aicf.in/api/players?name=%s&state=0&city=0", id)

	response, err := http.Get(url)
	if err != nil {
		return "", fmt.Errorf("Error in making the API call: %v", err)
	}

	defer response.Body.Close()

	var data map[string]interface{}
	if err := json.NewDecoder(response.Body).Decode(&data); err != nil {
		return "", fmt.Errorf("Error in decoding response body: %v", err)
	}

	dataArray, ok := data["data"].([]interface{})
	if !ok || len(dataArray) <= 0 {
		return "", fmt.Errorf("Error in getting dataArray: %v", err)
	}

	if len(dataArray) != 1 {
		return "", fmt.Errorf("Incorrect ID")
	}

	playerData, ok := dataArray[0].(map[string]interface{})
	if !ok {
		return "", fmt.Errorf("Error in getting playerData: %v", err)
	}

	aicfId, ok := playerData["aicf_id"].(string)
	if !ok {
		return "", fmt.Errorf("Error in getting AICF ID: %v", err)
	}

	return aicfId, nil
}

func main() {
	startTime := time.Now()

	var excelPath, sheetName string
	var tableRowNumber int
	fmt.Print("Enter excel file path with extension (no spaces anywhere in the file path): ")
	_, err := fmt.Scan(&excelPath)
	if err != nil {
		fmt.Println("Error in excel file name: ", err)
		return
	}

	fmt.Print("Enter sheet name: ")
	_, err = fmt.Scan(&sheetName)
	if err != nil {
		fmt.Println("Error in sheet name: ", err)
		return
	}

	fmt.Print("Enter row number where table starts: ")
	_, err = fmt.Scan(&tableRowNumber)
	if err != nil {
		fmt.Println("Error in table row number: ", err)
		return
	}

	xl, err := excelize.OpenFile(excelPath)
	if err != nil {
		fmt.Println("Error opening file: ", err)
		return
	}

	defer xl.Close()

	sheetDimension, err := xl.GetSheetDimension(sheetName)
	if err != nil {
		fmt.Println("Error getting sheet dimension: ", err)
		return
	}

	lastCol := strings.Split(sheetDimension, ":")[1][0]

	membershipStatusCell := string(lastCol+1) + strconv.Itoa(tableRowNumber)
	mcaIdCell := string(lastCol+2) + strconv.Itoa(tableRowNumber)

	err = xl.SetCellStr(sheetName, membershipStatusCell, "Membership_Status")
	if err != nil {
		fmt.Println("Error setting the Membership_Status column: ", err)
		return
	}

	err = xl.SetCellStr(sheetName, mcaIdCell, "MCA ID")
	if err != nil {
		fmt.Println("Error setting the MCA ID column: ", err)
		return
	}

	err = xl.Save()
	if err != nil {
		fmt.Println("Error saving excel file: ", err)
		return
	}

	membershipStatusColumn := string(lastCol + 1)
	mcaIdColumn := string(lastCol + 2)

	var aicfColumn, fideColumn string

	for colNum := byte(0); (colNum + 65) <= lastCol; colNum++ {
		colName := string(colNum + 65)
		cellNum := colName + strconv.Itoa(tableRowNumber)
		cellValue, err := xl.GetCellValue(sheetName, cellNum)
		if err != nil {
			fmt.Println("Error reading column cell value: ", err)
			return
		}

		if cellValue == "AICF ID" {
			aicfColumn = colName
		} else if cellValue == "FIDE ID" {
			fideColumn = colName
		}
	}

	for dataRowNumber := tableRowNumber + 1; ; dataRowNumber++ {
		firstCellValue, err := xl.GetCellValue(sheetName, "A" + strconv.Itoa(dataRowNumber))
		if err != nil {
			fmt.Println("Error fetching first cell value to verify the table end: ", err)
			return
		}

		if firstCellValue == "" {
			break
		}

		aicfId, err := xl.GetCellValue(sheetName, aicfColumn + strconv.Itoa(dataRowNumber))
		if err != nil {
			fmt.Println("Error getting cell value: ", err)
			return
		}

		fideId, err := xl.GetCellValue(sheetName, fideColumn + strconv.Itoa(dataRowNumber))
		if err != nil {
			fmt.Println("Error getting cell value: ", err)
			return
		}

		var membershipStatus bool

		if aicfId == "" && fideId == "" {
			membershipStatus = false
			err = xl.SetCellStr(sheetName, membershipStatusColumn + strconv.Itoa(dataRowNumber), "Check Manually")
			if err != nil {
				fmt.Println("Error setting membership status value: ", err)
				return
			}
			continue
		}

		if aicfId == "" {
			aicfId, err = getAICFId(fideId)
			if err != nil {
				fmt.Printf("Error setting AICF ID from FIDE ID: %v\n", fideId)
			}

			err = xl.SetCellStr(sheetName, aicfColumn + strconv.Itoa(dataRowNumber), aicfId)
			if err != nil {
				fmt.Println("Error setting AICF ID value: ", err)
				return
			}
		}

		mcaId, err := getMCAId(aicfId)
		if err != nil {
			fmt.Printf("ID: %v; Error getting MCA ID for AICF ID: %v\n", aicfId, err)
		} else {
			err = xl.SetCellStr(sheetName, mcaIdColumn + strconv.Itoa(dataRowNumber), mcaId)
			if err != nil {
				fmt.Println("Error setting MCA ID in sheet")
				return
			}
		}

		membershipStatusFIDE, err := getMembership(fideId)
		if err != nil {
			fmt.Printf("ID: %v; Error in getting membership status response FIDE ID: %v\n", fideId, err)
		}

		membershipStatusAICF, err := getMembership(aicfId)
		if err != nil {
			fmt.Printf("ID: %v; Error in getting membership status response AICF ID: %v\n", aicfId, err)
		}

		membershipStatus = membershipStatusFIDE || membershipStatusAICF

		if membershipStatus == false {
			err = xl.SetCellStr(sheetName, membershipStatusColumn + strconv.Itoa(dataRowNumber), "Check Manually")
			if err != nil {
				fmt.Println("Error setting membership status value: ", err)
				return
			}
		} else {
			err = xl.SetCellStr(sheetName, membershipStatusColumn + strconv.Itoa(dataRowNumber), "Yes")
			if err != nil {
				fmt.Println("Error setting membership status value: ", err)
				return
			}
		}

		time.Sleep(2 * time.Second)
	}

	err = xl.Save()
	if err != nil {
		fmt.Println("Error saving excel file: ", err)
		return
	}
	
	endTime := time.Now()

	fmt.Println("Time taken: ", endTime.Sub(startTime))

	os.Exit(0)
}
