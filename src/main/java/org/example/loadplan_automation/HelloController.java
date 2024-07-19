package org.example.loadplan_automation;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.loadplan_automation.Models.LotSummary_Model;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;
import java.util.stream.Collectors;

public class HelloController implements Initializable {
    private final ObservableList<Region_model> list = FXCollections.observableArrayList();
    private final ObservableList<Region_model> filteredList = FXCollections.observableArrayList();
    private final Set<String> regions = new HashSet<>();
    private final Set<String> allDivisions = new HashSet<>();
    private final Set<String> allGradeLevels = new HashSet<>();
    private final LinkedHashSet<Integer> selectedSchoolIDs = new LinkedHashSet<>();
    private final List<Integer> selectedCheckBoxValues = new ArrayList<>();
    private final List<LotSummary_Model> lotRowsModel =  new ArrayList<>();
    private int checkLotCount = 1;

    // For checkboxes of LOT numbers
    @FXML
    private CheckBox checkBoxLot10;
    @FXML
    private CheckBox checkBoxLot11;
    @FXML
    private CheckBox checkBoxLot12;
    @FXML
    private CheckBox checkBoxLot14;
    @FXML
    private CheckBox checkBoxLot15;
    @FXML
    private CheckBox checkBoxLot16;
    @FXML
    private CheckBox checkBoxLot17;
    @FXML
    private CheckBox checkBoxLot18;
    @FXML
    private CheckBox checkBoxLot19;

    // For table controls
    @FXML
    private ComboBox<String> drop_region;
    @FXML
    private ComboBox<String> drop_division;
    @FXML
    private ComboBox<String> drop_grade;
    @FXML
    private TextField searchSchoolID;
    @FXML
    private Button generate;
    @FXML
    private TableView<Region_model> table;
    @FXML
    private TableColumn<Region_model, Boolean> selectColumn;
    @FXML
    private CheckBox selectAllCheckbox;
    @FXML
    private TableColumn<Region_model, String> division;
    @FXML
    private TableColumn<Region_model, String> gradeLevel;
    @FXML
    private TableColumn<Region_model, String> region;
    @FXML
    private TableColumn<Region_model, Integer> schoolID;
    @FXML
    private TableColumn<Region_model, String> schoolName;

    @Override
    public void initialize(URL url, ResourceBundle resources) {
        initializeTableColumns();
        fetchDataFromPHP();
        insertTemplateLotRows();

        // Event listener for drop_region ComboBox
        drop_region.setOnAction(event -> {
            String selectedRegion = drop_region.getValue();
            System.out.println("Selected Region: " + selectedRegion);
            filterList();
            updateDivisionComboBox(selectedRegion);
            updateGradeComboBox(null);
        });

        // Event listener for drop_division ComboBox
        drop_division.setOnAction(event -> {
            String selectedDivision = drop_division.getValue();
            System.out.println("Selected Division: " + selectedDivision);
            filterList();
            updateGradeComboBox(selectedDivision);
        });

        // Event listener for drop_grade ComboBox
        drop_grade.setOnAction(event -> {
            String selectedGrade = drop_grade.getValue();
            System.out.println("Selected Grade: " + selectedGrade);
            filterList();
        });

        // Add listener to selectAllCheckbox
        selectAllCheckbox.selectedProperty().addListener((obs, wasSelected, isSelected) -> {
            for (Region_model item : filteredList) {
                item.getSelect().setSelected(isSelected);
                updateSelectedSchoolIDs(item, isSelected);
            }
        });

        // Event handler for the generate button
        generate.setOnAction(event -> {
            System.out.println("Selected School IDs: " + selectedSchoolIDs);
            readExcelFile(selectedSchoolIDs);
        });

        // Listener for searchSchoolID TextField
        searchSchoolID.textProperty().addListener((observable, oldValue, newValue) -> {
            filterList();
        });

        // Add listeners to checkboxes
        addCheckboxListeners();
    }

    private void insertTemplateLotRows() {
        lotRowsModel.add(new LotSummary_Model(12, 53, 27)); // For lot 12
        lotRowsModel.add(new LotSummary_Model(14, 81, 18)); // For lot 14
        lotRowsModel.add(new LotSummary_Model(16, 115, 16)); // For lot 16
        lotRowsModel.add(new LotSummary_Model(17, 132, 12)); // For lot 17
        lotRowsModel.add(new LotSummary_Model(18, 145, 8)); // For lot 18
    }

    private void initializeTableColumns() {
        // Initialize table columns with property values
        selectColumn.setCellValueFactory(new PropertyValueFactory<>("select"));
        region.setCellValueFactory(new PropertyValueFactory<>("region"));
        division.setCellValueFactory(new PropertyValueFactory<>("division"));
        schoolID.setCellValueFactory(new PropertyValueFactory<>("schoolID"));
        schoolName.setCellValueFactory(new PropertyValueFactory<>("schoolName"));
        gradeLevel.setCellValueFactory(new PropertyValueFactory<>("gradeLevel"));

        table.setItems(filteredList);

        // Add listeners to rows to update selected school IDs
        table.setRowFactory(tv -> {
            TableRow<Region_model> row = new TableRow<>();
            row.itemProperty().addListener((obs, previousItem, currentItem) -> {
                if (previousItem != null) {
                    previousItem.getSelect().selectedProperty().removeListener((obs1, wasSelected, isSelected) -> updateSelectedSchoolIDs(previousItem, isSelected));
                }
                if (currentItem != null) {
                    currentItem.getSelect().selectedProperty().addListener((obs1, wasSelected, isSelected) -> updateSelectedSchoolIDs(currentItem, isSelected));
                }
            });
            return row;
        });
    }

    private void updateSelectedSchoolIDs(Region_model regionModel, boolean isSelected) {
        if (isSelected) {
            selectedSchoolIDs.add(regionModel.getSchoolID());
        } else {
            selectedSchoolIDs.remove(regionModel.getSchoolID());
        }
    }

    private void addCheckboxListeners() {
        CheckBox[] checkBoxes = {checkBoxLot10, checkBoxLot11, checkBoxLot12, checkBoxLot14, checkBoxLot15,
                checkBoxLot16, checkBoxLot17, checkBoxLot18, checkBoxLot19};

        for (CheckBox checkBox : checkBoxes) {
            checkBox.selectedProperty().addListener((obs, wasSelected, isSelected) -> {
                int checkboxValue = Integer.parseInt(checkBox.getText());
                if (isSelected) {
                    selectedCheckBoxValues.add(checkboxValue);
                    checkLotCount++;
                    if (checkLotCount > 5) {
                        for (CheckBox cb : checkBoxes) {
                            if (!cb.isSelected()) {
                                cb.setDisable(true);
                            }
                        }
                    }
                } else {
                    selectedCheckBoxValues.remove((Integer) checkboxValue);
                    checkLotCount--;
                    if (checkLotCount <= 5) {
                        for (CheckBox cb : checkBoxes) {
                            cb.setDisable(false);
                        }
                    }
                }
                // Sort the list of selected checkbox values
                Collections.sort(selectedCheckBoxValues);
                System.out.println("Selected Checkboxes Values (sorted): " + selectedCheckBoxValues);
            });
        }
    }

    private void fetchDataFromPHP() {
        new Thread(() -> {
            try {
                // Construct URL for PHP data
                String urlString = "http://192.168.1.229/NIKKA/get_data.php";
                URL url = new URL(urlString);
                HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                conn.setRequestMethod("GET");

                // Read data from URL
                BufferedReader in = new BufferedReader(new InputStreamReader(conn.getInputStream()));
                StringBuilder content = new StringBuilder();
                String inputLine;
                while ((inputLine = in.readLine()) != null) {
                    content.append(inputLine);
                }
                in.close();
                conn.disconnect();

                // Parse JSON data
                JSONObject json = new JSONObject(content.toString());

                // Map to keep track of schoolID and Region_model
                Map<Integer, Region_model> schoolMap = new HashMap<>();

                for (String tableName : json.keySet()) {
                    JSONArray tableData = json.getJSONArray(tableName);
                    for (int i = 0; i < tableData.length(); i++) {
                        JSONObject row = tableData.getJSONObject(i);
                        int schoolID = row.getInt("SchoolID");
                        String gradeLevel = row.getString("GradeLevel");

                        if (schoolMap.containsKey(schoolID)) {
                            // If schoolID already exists, modify the existing Region_model
                            Region_model existingModel = schoolMap.get(schoolID);
                            String existingGradeLevel = existingModel.getGradeLevel();
                            String newGradeLevel = existingGradeLevel + "," + gradeLevel;
                            existingModel.setGradeLevel(newGradeLevel);
                        } else {
                            // If schoolID does not exist, create a new Region_model
                            Region_model regionModel = new Region_model(
                                    row.getString("Region"),
                                    row.getString("Division"),
                                    schoolID,
                                    row.getString("School Name"),
                                    gradeLevel
                            );
                            list.add(regionModel);
                            schoolMap.put(schoolID, regionModel);
                        }

                        // Collect other data as needed
                        regions.add(row.getString("Region"));
                        allDivisions.add(row.getString("Division"));
                        allGradeLevels.add(gradeLevel);
                    }
                }

                // Update ComboBoxes and TableView on JavaFX Application Thread
                Platform.runLater(() -> {
                    ObservableList<String> regionItems = FXCollections.observableArrayList(regions);
                    regionItems.add(0, "All");
                    drop_region.setItems(regionItems);
                    drop_region.setValue("All"); // Set default value to "All"

                    ObservableList<String> divisionItems = FXCollections.observableArrayList(allDivisions);
                    divisionItems.add(0, "All");
                    drop_division.setItems(divisionItems);
                    drop_division.setValue("All"); // Set default value to "All"

                    ObservableList<String> gradeItems = FXCollections.observableArrayList(allGradeLevels);
                    gradeItems.add(0, "All");
                    drop_grade.setItems(gradeItems);
                    drop_grade.setValue("All"); // Set default value to "All"

                    filterList(); // Initialize the table with all data
                });

            } catch (Exception e) {
                e.printStackTrace();
            }
        }).start();
    }

    private void filterList() {
        filteredList.setAll(list.stream()
                .filter(this::matchesFilter)
                .collect(Collectors.toList()));
    }

    private boolean matchesFilter(Region_model regionModel) {
        String selectedRegion = drop_region.getValue();
        String selectedDivision = drop_division.getValue();
        String selectedGrade = drop_grade.getValue();
        String searchText = searchSchoolID.getText().toLowerCase();

        boolean matchesRegion = selectedRegion == null || selectedRegion.equals("All") || regionModel.getRegion().equals(selectedRegion);
        boolean matchesDivision = selectedDivision == null || selectedDivision.equals("All") || regionModel.getDivision().equals(selectedDivision);
        boolean matchesGrade = selectedGrade == null || selectedGrade.equals("All") || Arrays.asList(regionModel.getGradeLevel().split(",")).contains(selectedGrade);
        boolean matchesSearch = searchText.isEmpty() || String.valueOf(regionModel.getSchoolID()).contains(searchText);

        return matchesRegion && matchesDivision && matchesGrade && matchesSearch;
    }

    private void updateDivisionComboBox(String selectedRegion) {
        Set<String> filteredDivisions = new HashSet<>();
        for (Region_model item : list) {
            if (selectedRegion.equals("All") || item.getRegion().equals(selectedRegion)) {
                filteredDivisions.add(item.getDivision());
            }
        }

        ObservableList<String> divisionItems = FXCollections.observableArrayList(filteredDivisions);
        divisionItems.add(0, "All");
        drop_division.setItems(divisionItems);
        drop_division.setValue("All"); // Set default value to "All"
    }

    private void updateGradeComboBox(String selectedDivision) {
        Set<String> filteredGrades = new HashSet<>();
        for (Region_model item : list) {
            if ((selectedDivision == null || selectedDivision.equals("All") || item.getDivision().equals(selectedDivision))
                    && (drop_region.getValue() == null || drop_region.getValue().equals("All") || item.getRegion().equals(drop_region.getValue()))) {
                filteredGrades.addAll(Arrays.asList(item.getGradeLevel().split(",")));
            }
        }

        ObservableList<String> gradeItems = FXCollections.observableArrayList(filteredGrades);
        gradeItems.add(0, "All");
        drop_grade.setItems(gradeItems);
        drop_grade.setValue("All"); // Set default value to "All"
    }

    public void readExcelFile(LinkedHashSet<Integer> selectedSchoolIDs) {
        // Paths to the source data file, template file, and output file
        String sourceData = "C:\\Users\\5CG6105SVT\\Desktop\\LTE-SM-2023-Allocation-List-Lots-1415161718-Nikka-Trading.xlsx";
        String templatePath = "res/template/Load_Plan_Template.xlsx";
        String outputPath = "C:\\Users\\5CG6105SVT\\Desktop\\output.xlsx";

        // Show warning if no School IDs are selected
        if (selectedSchoolIDs.isEmpty()) {
            Alert alert = new Alert(Alert.AlertType.WARNING, "No School IDs selected.", ButtonType.OK);
            alert.showAndWait();
            return;
        }

        // List to store sets of row numbers with the same school ID
        List<List<Integer>> sameSchoolIDRows = new ArrayList<>();

        try (FileInputStream sourceFis = new FileInputStream(sourceData);
             FileInputStream templateFis = new FileInputStream(templatePath);
             Workbook sourceWorkbook = new XSSFWorkbook(sourceFis);
             Workbook templateWorkbook = new XSSFWorkbook(templateFis)) {

            // Get the first sheet of the source data and template
            Sheet sourceSheet = sourceWorkbook.getSheetAt(0);
            Sheet templateSheet = templateWorkbook.getSheetAt(0);

            // Prepare the row and style for writing checkbox values to the template
            int startLotColumnIndex = 6; // Column G
            int LotRowIndex = 10; // Row 11 (0-based index)
            Row styleRow = templateSheet.getRow(12); // Row with styles in the template
            CellStyle[] templateStyles = new CellStyle[styleRow.getLastCellNum()];
            for (int i = 0; i < styleRow.getLastCellNum(); i++) {
                templateStyles[i] = styleRow.getCell(i).getCellStyle();
            }
            Row checkboxRow = templateSheet.getRow(LotRowIndex);
            if (checkboxRow == null) {
                checkboxRow = templateSheet.createRow(LotRowIndex);
            }

            // Set the selected LOT values in the template
            for (int i = 0; i < selectedCheckBoxValues.size() && i < 5; i++) {
                Cell cell = checkboxRow.createCell(startLotColumnIndex + i, CellType.NUMERIC);
                cell.setCellValue("LOT "+selectedCheckBoxValues.get(i));
                cell.setCellStyle(templateStyles[startLotColumnIndex + i]);


            }

            int outputRowNum = 12; // Starting row index in the output sheet
            int rowCount = 1;
            int maxRowCount = 23;
            String previousSchoolID = null;
            List<Integer> currentSet = new ArrayList<>();
            int listNum = 1;

            // Iterate over each selected School ID
            for (Integer selectedSchoolID : selectedSchoolIDs) {
                // Iterate over rows in the source sheet to find matching School IDs
                for (Row row : sourceSheet) {
                    Cell cell = row.getCell(2); // Assuming School ID is in the third column (index 2)
                    if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                        int schoolID = (int) cell.getNumericCellValue();
                        if (schoolID == selectedSchoolID) {

                            // Data from the source sheet
                            String cellDivision = getCellValue(row.getCell(1));
                            String cellSchoolID = getCellValue(row.getCell(2));
                            String cellSchoolName = getCellValue(row.getCell(3));
                            String cellGradeLevel = getCellValue(row.getCell(4));

                            // Check if the current School ID is the same as the previous one
                            if (previousSchoolID != null && previousSchoolID.equals(cellSchoolID)) {
                                currentSet.add(outputRowNum); // Add the row number to the current set
                            } else {
                                if (!currentSet.isEmpty()) {
                                    sameSchoolIDRows.add(new ArrayList<>(currentSet)); // Store the current set
                                    currentSet.clear(); // Clear the current set
                                }
                                currentSet.add(outputRowNum); // Start a new set with the current row number
                            }

                            // Create a new row in the template sheet
                            Row templateRow = templateSheet.createRow(outputRowNum++);

                            // Write values to the new row in the template sheet with styles
                            if (!cellSchoolID.equals(previousSchoolID)) {
                                createStyledCell(templateRow, 0, String.valueOf(listNum), templateStyles[0]);
                                listNum++;
                            }

                            // Update previousSchoolID to current cellSchoolID
                            previousSchoolID = cellSchoolID;

                            createStyledCell(templateRow, 1, cellDivision, templateStyles[1]);
                            createStyledCell(templateRow, 2, "", templateStyles[2]);
                            createStyledCell(templateRow, 3, cellSchoolID, templateStyles[3]);
                            createStyledCell(templateRow, 4, cellSchoolName, templateStyles[4]);
                            createStyledCell(templateRow, 5, cellGradeLevel, templateStyles[5]);
                            createStyledCell(templateRow, 6, "", templateStyles[6]);
                            createStyledCell(templateRow, 7, "", templateStyles[7]);
                            createStyledCell(templateRow, 8, "", templateStyles[8]);
                            createStyledCell(templateRow, 9, "", templateStyles[9]);
                            createStyledCell(templateRow, 10, "", templateStyles[10]);
                            createStyledCell(templateRow, 11, "", templateStyles[11]);
                            createStyledCell(templateRow, 12, "", templateStyles[12]);
                            createStyledCell(templateRow, 13, "", templateStyles[13]);
                            createStyledCell(templateRow, 14, "", templateStyles[14]);

                            // Shift rows down if max row count is reached
                            if (rowCount == maxRowCount) {
                                int rowIndex = 37; // Insert before row 38
                                templateSheet.shiftRows(rowIndex, templateSheet.getLastRowNum(), 35);
                            }
                            rowCount++;
                        }
                    }
                }

                // Store the last set of rows if any
                if (!currentSet.isEmpty()) {
                    sameSchoolIDRows.add(new ArrayList<>(currentSet));
                    currentSet.clear();
                }

                // Reset previousSchoolID for the next selectedSchoolID
                previousSchoolID = null;
            }

            // Merge cells in columns A to D for each set of row numbers with the same School ID
            for (List<Integer> rowSet : sameSchoolIDRows) {
                if (rowSet.size() > 1) {
                    for (int colNum = 0; colNum < 5; colNum++) {
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowSet.get(0), rowSet.get(rowSet.size() - 1), colNum, colNum);
                        templateSheet.addMergedRegion(cellRangeAddress);
                    }
                }
            }

            // Write the changes to the output file
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                templateWorkbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Method to get cell value as a string
    private static String getCellValue(Cell cell) {
        if (cell != null) {
            return switch (cell.getCellType()) {
                case STRING -> cell.getStringCellValue();
                case NUMERIC -> String.valueOf(cell.getNumericCellValue());
                case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                default -> "Unknown Cell Type";
            };
        }
        return "Unknown Cell Type";
    }

    // Method to create a styled cell
    private static void createStyledCell(Row row, int column, String value, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }
}

