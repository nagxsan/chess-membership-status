# Chess Membership Status
I am an avid chess player and my father (IA Sandesh Nagarnaik) has been a Chief Arbiter in many chess tournaments. I have always seen him spending a lot of time verifying the players participating in the tournament. The general process is always to look up the player's AICF/FIDE ID on the website and check the "Membership Status" of the player whether it is "Active/Not Active". I have done the same thing for the MCA (Maharashtra) membership as well I figured that if I could automate this process by making an API call to the website and parsing the response generated, then it may save him the time of verifying the players manually.

## Pre-requisites (IMPORTANT)
Since I have solved for this specific use case, there are a few pointers that you need to follow if this application needs to be executed error free.  
1. The players' information and everything must be in a **'.xlsx'** file only (try having no spaces in the file-name).
2. The columns containing the FIDE ID's or the AICF ID's must be named **FIDE ID** or **AICF ID** respectively.
3. The column containing the MCA ID must be named **MCA ID**
4. Ensure that there is at least one completely empty row after the player's details table in the excel file.
5. This is not compulsory but it may avoid unpredictable errors (if any). Keep the excel file which you need to evaluate and the executable application generated in the same folder.
6. Make sure that the number of columns in the table are less than 26 that is from (A-Y) at max. If that's not the case, you can just create a different excel file with only the ID's which adhere to these instructions and execute the application.
7. Ensure that the excel file is closed when the application is executing.
8. Please answer the required details when the program starts accurately, and then sit back and relax since the application may take a while to run!

## How to use
### Step 1: Clone the repository
You will need to clone the repository to your local device:  
```
git clone https://github.com/daredevil0905/chess-membership-status.git
```

### Step 2: Install go and its dependencies
You will need to install `go` and its dependencies on your local machine to run the code. You can download go from the official website https://go.dev/dl/

### Step 3: Navigate to the Project Directory
Navigate to the cloned project directory in your device.  
```
cd chess-membership-status
```

### Step 4: Install the project dependencies
Run the following command to install project dependencies in your terminal. This will install the dependencies listed in the `go.mod` file.  
```
go mod download
```

### Step 5: Build the application
Now you need to build the application. For doing this execute the following code in the terminal. This will build the code in `main.go` and generate an executable file that you can directly run on your device.  
```
go build
```

### Step 6: Run the application
Run the executable file by double-clicking it or in the terminal with the following command (you may need to specify relative path by using './' before executable file's name, choose whatever works for you):  
```
chess-membership-status.exe
```
