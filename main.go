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

func getNameFromAICFOrFIDEId(id string) (string, error) {
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

  var firstName, middleName, lastName string
  if playerData["first_name"] != nil {
    firstName = playerData["first_name"].(string)
  }

  if playerData["middle_name"] != nil {
    middleName = playerData["middle_name"].(string)
  }

  if playerData["last_name"] != nil {
    lastName = playerData["last_name"].(string)
  }

  name := firstName + " " + middleName + " " + lastName

  return name, nil
}

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
	if len(id) < 8 || !strings.Contains(strings.ToLower(id), "mh") {
		return "", fmt.Errorf("error: incorrect AICF ID")
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
		return "", fmt.Errorf("error in creating new request: %v", err)
	}

	for key, value := range headers {
		req.Header.Set(key, value)
	}

	response, err := client.Do(req)
	if err != nil {
		return "", fmt.Errorf("error in making an API call: %v", err)
	}
	defer response.Body.Close()

	if response.StatusCode != http.StatusOK {
		return "", fmt.Errorf("unexpected status code; API call failed")
	}

	var reader io.ReadCloser
	switch response.Header.Get("Content-Encoding") {
	case "gzip":
		reader, err = gzip.NewReader(response.Body)
		if err != nil {
			return "", fmt.Errorf("error in unzipping response body: %v", err)
		}
		defer reader.Close()
	default:
		reader = response.Body
	}

	body, err := io.ReadAll(reader)
	if err != nil {
		return "", fmt.Errorf("error in reading response body: %v", err)
	}

	responseHTML := string(body)
  
  replaceSpacesRe := regexp.MustCompile(`\s+>`)
  responseHTML = replaceSpacesRe.ReplaceAllString(responseHTML, ">")

  var (
    tableHeadings,
    tableData [] string
  )

  var matches [][]string

  tableHeadingsRe := regexp.MustCompile(`<th>(.*?)</th>`)
  matches = tableHeadingsRe.FindAllStringSubmatch(responseHTML, -1)
  for _, match := range matches {
    tableHeadings = append(tableHeadings, match[1])
  }
  
  numCols := len(tableHeadings)

  tableDataRe := regexp.MustCompile(`<td>(.*?)</td>`)
  matches = tableDataRe.FindAllStringSubmatch(responseHTML, -1)
  for _, match := range matches {
    tableData = append(tableData, match[1])
  }

  var mcaLicenses [][]string
  var mcaLicense []string
  for idx, data := range tableData {
    if idx % numCols == 1 || idx % numCols == 2 || idx % numCols == numCols - 1 {
      if idx % numCols == 1 {
        nameRe := regexp.MustCompile(`<strong>(.*?)</strong>`)
        matches = nameRe.FindAllStringSubmatch(data, -1)
        match := strings.Split(matches[0][1], " ")
        name := match[0] + " " + match[len(match) - 1]
        data = name
      }
      if idx % numCols == 2 {
        data = data[:len(data) - 3]
      }
      mcaLicense = append(mcaLicense, data)
    }
    if idx % numCols == numCols - 1 {
      mcaLicenses = append(mcaLicenses, mcaLicense)
      mcaLicense = []string{}
    }
  }

  if len(mcaLicenses) == 0 {
    return "", fmt.Errorf("%s: No MCA player record found.", id)
  }

  for _, license := range mcaLicenses {
    if license[2] == "" {
      continue
    }
    if license[1] != "" {
      return license[0] + ";" + license[1], nil
    }
  }
  return "", fmt.Errorf("%s: no MCA ID found", id)
}

func getAICFId(id string) (string, error) {
	url := fmt.Sprintf("https://admin.aicf.in/api/players?name=%s&state=0&city=0", id)

	response, err := http.Get(url)
	if err != nil {
		return "", fmt.Errorf("error in making the API call: %v", err)
	}

	defer response.Body.Close()

	var data map[string]interface{}
	if err := json.NewDecoder(response.Body).Decode(&data); err != nil {
		return "", fmt.Errorf("error in decoding response body: %v", err)
	}

	dataArray, ok := data["data"].([]interface{})
	if !ok || len(dataArray) <= 0 {
		return "", fmt.Errorf("error in getting dataArray: %v", err)
	}

	if len(dataArray) != 1 {
		return "", fmt.Errorf("incorrect ID")
	}

	playerData, ok := dataArray[0].(map[string]interface{})
	if !ok {
		return "", fmt.Errorf("error in getting playerData: %v", err)
	}

	aicfId, ok := playerData["aicf_id"].(string)
	if !ok {
		return "", fmt.Errorf("error in getting AICF ID: %v", err)
	}

	return aicfId, nil
}

func main() {
	startTime := time.Now()

	fmt.Println("Created and Developed by Sanchet Nagarnaik")

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
  membershipNameCell := string(lastCol+3) + strconv.Itoa(tableRowNumber)
  mcaNameCell := string(lastCol+4) + strconv.Itoa(tableRowNumber)

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
  
  err = xl.SetCellStr(sheetName, membershipNameCell, "Membership Name")
	if err != nil {
		fmt.Println("Error setting the Membership Name column: ", err)
		return
	}

  err = xl.SetCellStr(sheetName, mcaNameCell, "MCA Name")
	if err != nil {
		fmt.Println("Error setting the MCA Name column: ", err)
		return
	}

	err = xl.Save()
	if err != nil {
		fmt.Println("Error saving excel file: ", err)
		return
	}

	membershipStatusColumn := string(lastCol + 1)
	mcaIdColumn := string(lastCol + 2)
  membershipNameColumn := string(lastCol + 3)
  mcaNameColumn := string(lastCol + 4)

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
		firstCellValue, err := xl.GetCellValue(sheetName, "A"+strconv.Itoa(dataRowNumber))
		if err != nil {
			fmt.Println("Error fetching first cell value to verify the table end: ", err)
			return
		}

		if firstCellValue == "" {
			break
		}

		aicfId, err := xl.GetCellValue(sheetName, aicfColumn+strconv.Itoa(dataRowNumber))
		if err != nil {
			fmt.Println("Error getting cell value: ", err)
			return
		}

		fideId, err := xl.GetCellValue(sheetName, fideColumn+strconv.Itoa(dataRowNumber))
		if err != nil {
			fmt.Println("Error getting cell value: ", err)
			return
		}

		var membershipStatus bool

		if aicfId == "" && fideId == "" {
			membershipStatus = false
			err = xl.SetCellStr(sheetName, membershipStatusColumn+strconv.Itoa(dataRowNumber), "Check Manually")
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

			err = xl.SetCellStr(sheetName, aicfColumn+strconv.Itoa(dataRowNumber), aicfId)
			if err != nil {
				fmt.Println("Error setting AICF ID value: ", err)
				return
			}
		}

		mcaDetails, err := getMCAId(aicfId)
		if err != nil {
			fmt.Printf("ID: %v; Error getting MCA Details for AICF ID: %v\n", aicfId, err)
		} else {
      details := strings.Split(mcaDetails, ";")
      mcaName := details[0]
      mcaId := details[1]

			err = xl.SetCellStr(sheetName, mcaIdColumn+strconv.Itoa(dataRowNumber), mcaId)
			if err != nil {
				fmt.Println("Error setting MCA ID in sheet")
				return
			}

      err = xl.SetCellStr(sheetName, mcaNameColumn+strconv.Itoa(dataRowNumber), mcaName)
			if err != nil {
				fmt.Println("Error setting MCA Name in sheet")
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

		if !membershipStatus {
			err = xl.SetCellStr(sheetName, membershipStatusColumn+strconv.Itoa(dataRowNumber), "Check Manually")
			if err != nil {
				fmt.Println("Error setting membership status value: ", err)
				return
			}
		} else {
			err = xl.SetCellStr(sheetName, membershipStatusColumn+strconv.Itoa(dataRowNumber), "Active")
			if err != nil {
				fmt.Println("Error setting membership status value: ", err)
				return
			}
		}

    membershipNameFIDE, err := getNameFromAICFOrFIDEId(fideId)
    if err != nil {
      fmt.Printf("ID: %v; Error in getting membership name from FIDE ID: %v\n", fideId, err)
    }

    membershipNameAICF, err := getNameFromAICFOrFIDEId(aicfId)
    if err != nil {
      fmt.Printf("ID: %v; Error in getting membership name from AICF ID: %v\n", aicfId, err)
    }

    var membershipName string
    if membershipNameFIDE != "" {
      membershipName = membershipNameFIDE
    } else {
      membershipName = membershipNameAICF
    }

    err = xl.SetCellStr(sheetName, membershipNameColumn+strconv.Itoa(dataRowNumber), membershipName)
		if err != nil {
			fmt.Println("Error setting membership name value: ", err)
			return
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
