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

    @FXML
    private ComboBox<String> drop_region;
    @FXML
    private ComboBox<String> drop_division;
    @FXML
    private ComboBox<String> drop_grade;
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
    }

    private void initializeTableColumns() {
        selectColumn.setCellValueFactory(new PropertyValueFactory<>("select"));
        region.setCellValueFactory(new PropertyValueFactory<>("region"));
        division.setCellValueFactory(new PropertyValueFactory<>("division"));
        schoolID.setCellValueFactory(new PropertyValueFactory<>("schoolID"));
        schoolName.setCellValueFactory(new PropertyValueFactory<>("schoolName"));
        gradeLevel.setCellValueFactory(new PropertyValueFactory<>("gradeLevel"));

        table.setItems(filteredList);

        // Add listeners to checkboxes
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

    private void fetchDataFromPHP() {
        new Thread(() -> {
            try {
                // Construct URL with query parameter
                String urlString = "http://192.168.254.104/NIKKA/get_data.php";
                URL url = new URL(urlString);
                HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                conn.setRequestMethod("GET");

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
                // Handle error gracefully, e.g., by showing an alert to the user
            }
        }).start();
    }


    private void filterList() {
        String selectedRegion = drop_region.getValue();
        String selectedDivision = drop_division.getValue();
        String selectedGrade = drop_grade.getValue();

        filteredList.setAll(list.stream()
                .filter(regionModel -> "All".equals(selectedRegion) || regionModel.getRegion().equals(selectedRegion))
                .filter(regionModel -> "All".equals(selectedDivision) || regionModel.getDivision().equals(selectedDivision))
                .filter(regionModel -> "All".equals(selectedGrade) || regionModel.getGradeLevel().equals(selectedGrade))
                .collect(Collectors.toList()));
    }

    private void updateDivisionComboBox(String regionName) {
        Set<String> divisions = new HashSet<>();
        if ("All".equals(regionName)) {
            divisions = allDivisions;
        } else {
            divisions = list.stream()
                    .filter(regionModel -> regionModel.getRegion().equals(regionName))
                    .map(Region_model::getDivision)
                    .collect(Collectors.toSet());
        }

        ObservableList<String> divisionItems = FXCollections.observableArrayList(divisions);
        divisionItems.add(0, "All");
        drop_division.setItems(divisionItems);
        drop_division.setValue("All"); // Set default value to "All"
    }

    private void updateGradeComboBox(String divisionName) {
        Set<String> gradeLevels = new HashSet<>();
        if ("All".equals(divisionName) || divisionName == null) {
            gradeLevels = allGradeLevels;
        } else {
            gradeLevels = list.stream()
                    .filter(regionModel -> regionModel.getDivision().equals(divisionName))
                    .map(Region_model::getGradeLevel)
                    .collect(Collectors.toSet());
        }

        ObservableList<String> gradeItems = FXCollections.observableArrayList(gradeLevels);
        gradeItems.add(0, "All");
        drop_grade.setItems(gradeItems);
        drop_grade.setValue("All"); // Set default value to "All"
    }

    public static void readExcelFile(LinkedHashSet<Integer> selectedSchoolIDs) {
        String sourceData = "C:\\Users\\5CG6105SVT\\Desktop\\LTE-SM-2023-Allocation-List-Lots-1415161718-Nikka-Trading.xlsx";
        String templatePath = "res/template/Load_Plan_Template.xlsx";
        String outputPath = "C:\\Users\\5CG6105SVT\\Desktop\\output.xlsx";

        List<List<Integer>> sameSchoolIDRows = new ArrayList<>(); // List to store sets of row numbers with same school ID

        try (FileInputStream sourceFis = new FileInputStream(sourceData);
             FileInputStream templateFis = new FileInputStream(templatePath);
             Workbook sourceWorkbook = new XSSFWorkbook(sourceFis);
             Workbook templateWorkbook = new XSSFWorkbook(templateFis)) {

            // Get the first sheet of the source data
            Sheet sourceSheet = sourceWorkbook.getSheetAt(0);

            // Get the first sheet of the template
            Sheet templateSheet = templateWorkbook.getSheetAt(0);

            // Get the style of row 13 of the template
            Row styleRow = templateSheet.getRow(12);
            CellStyle[] templateStyles = new CellStyle[styleRow.getLastCellNum()];
            for (int i = 0; i < styleRow.getLastCellNum(); i++) {
                templateStyles[i] = styleRow.getCell(i).getCellStyle();
            }

            int outputRowNum = 12; // Starting row index (zero-based) in the output sheet
            String previousSchoolID = null; // Variable to store previous school ID

            List<Integer> currentSet = new ArrayList<>(); // Current set of rows with the same school ID

            int listNum = 1;
            // Iterate over each school ID in the selectedSchoolIDs set
            for (Integer selectedSchoolID : selectedSchoolIDs) {
                // Iterate over the rows in the source sheet and find matching school IDs
                for (Row row : sourceSheet) {
                    Cell cell = row.getCell(2); // Assuming school ID is in the third column (index 2)
                    if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                        int schoolID = (int) cell.getNumericCellValue();
                        if (schoolID == selectedSchoolID) {

                            String cellDivision = getCellValue(row.getCell(1));
                            String cellSchoolID = getCellValue(row.getCell(2));
                            String cellSchoolName = getCellValue(row.getCell(3));
                            String cellGradeLevel = getCellValue(row.getCell(4));

                            // Check if the current school ID is the same as the previous one
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

            // Merge cells in columns A to D for each set of row numbers with the same school ID
            for (List<Integer> rowSet : sameSchoolIDRows) {
                if (rowSet.size() > 1) {
                    for (int colNum = 0; colNum < 5; colNum++) {
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowSet.get(0), rowSet.get(rowSet.size() - 1), colNum, colNum);
                        templateSheet.addMergedRegion(cellRangeAddress);
                    }
                }
            }

            // Write the changes to a new file
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                templateWorkbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

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

    private static void createStyledCell(Row row, int column, String value, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

}
