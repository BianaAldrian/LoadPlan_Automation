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
import org.example.loadplan_automation.Models.GradeLevelSets_Model;
import org.example.loadplan_automation.Models.ItemSets_Model;
import org.example.loadplan_automation.Models.LotSummary_Model;
import org.example.loadplan_automation.Models.SummaryItems_Model;
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

    // Template for summary
    private void insertTemplateLotRows() {
        lotRowsModel.add(new LotSummary_Model(10, 2, 33)); // For lot 10
        lotRowsModel.add(new LotSummary_Model(11, 35, 51)); // For lot 11
        lotRowsModel.add(new LotSummary_Model(12, 53, 79)); // For lot 12
        lotRowsModel.add(new LotSummary_Model(14, 81, 98)); // For lot 14
        lotRowsModel.add(new LotSummary_Model(15, 100, 113)); // For lot 15
        lotRowsModel.add(new LotSummary_Model(16, 115, 130)); // For lot 16
        lotRowsModel.add(new LotSummary_Model(17, 132, 143)); // For lot 17
        lotRowsModel.add(new LotSummary_Model(18, 145, 152)); // For lot 18
        lotRowsModel.add(new LotSummary_Model(19, 154, 203)); // For lot 18
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
                String urlString = "http://192.168.254.100/NIKKA/get_data.php";
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
        String sourceData = "res/Allocation Master List.xlsx";
        String templatePath = "res/template/Load_Plan_Template.xlsx";
        String summaryTemplate = "res/template/Allocation_Summary_Template.xlsx";
        String itemBases = "res/bases.xlsx";
        String outputPath = "C:\\Users\\5CG6105SVT\\Desktop\\output.xlsx";

        if (selectedSchoolIDs.isEmpty()) {
            Alert alert = new Alert(Alert.AlertType.WARNING, "No School IDs selected.", ButtonType.OK);
            alert.showAndWait();
            return;
        }

        List<List<Integer>> sameSchoolIDRows = new ArrayList<>();

        try (FileInputStream sourceFis = new FileInputStream(sourceData);
             FileInputStream templateFis = new FileInputStream(templatePath);
             FileInputStream summaryTemplateFis = new FileInputStream(summaryTemplate);
             FileInputStream itemBasesFis = new FileInputStream(itemBases);
             Workbook sourceWorkbook = new XSSFWorkbook(sourceFis);
             Workbook templateWorkbook = new XSSFWorkbook(templateFis);
             Workbook summaryTemplateWorkbook = new XSSFWorkbook(summaryTemplateFis);
             Workbook itemBasesWorkbook = new XSSFWorkbook(itemBasesFis)) {

            Sheet sourceSheet = sourceWorkbook.getSheetAt(0);
            Sheet templateSheet = templateWorkbook.getSheetAt(0);
            Sheet summaryTemplateSheet = summaryTemplateWorkbook.getSheetAt(0);

            int startLotColumnIndex = 6;
            int lotRowIndex = 10;
            Row styleRow = templateSheet.getRow(12);
            CellStyle[] templateStyles = new CellStyle[styleRow.getLastCellNum()];
            for (int i = 0; i < styleRow.getLastCellNum(); i++) {
                templateStyles[i] = styleRow.getCell(i).getCellStyle();
            }
            Row checkboxRow = templateSheet.getRow(lotRowIndex);
            if (checkboxRow == null) {
                checkboxRow = templateSheet.createRow(lotRowIndex);
            }

            for (int i = 0; i < selectedCheckBoxValues.size() && i < 5; i++) {
                Cell cell = checkboxRow.createCell(startLotColumnIndex + i, CellType.STRING);
                cell.setCellValue("LOT " + selectedCheckBoxValues.get(i));
                cell.setCellStyle(templateStyles[startLotColumnIndex + i]);
            }

            int outputRowNum = 12;
            int rowCount = 1;
            int maxRowCount = 23;
            String previousSchoolID = null;
            List<Integer> currentSet = new ArrayList<>();
            int listNum = 1;

            for (Integer selectedSchoolID : selectedSchoolIDs) {
                for (Row row : sourceSheet) {
                    Cell cell = row.getCell(2);
                    if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                        int schoolID = (int) cell.getNumericCellValue();
                        if (schoolID == selectedSchoolID) {
                            String cellDivision = getCellValue(row.getCell(1));
                            String cellSchoolID = getCellValue(row.getCell(2));
                            String cellSchoolName = getCellValue(row.getCell(3));
                            String cellGradeLevel = getCellValue(row.getCell(4));

                            List<Integer> setOfItem = new ArrayList<>();
                            for (int i = 0; i < selectedCheckBoxValues.size() && i < 5; i++) {
                                findHeaderColumnRange(sourceSheet, itemBasesWorkbook, "LOT " + selectedCheckBoxValues.get(i), row.getCell(2), setOfItem);
                            }

                            if (previousSchoolID != null && previousSchoolID.equals(cellSchoolID)) {
                                currentSet.add(outputRowNum);
                            } else {
                                if (!currentSet.isEmpty()) {
                                    sameSchoolIDRows.add(new ArrayList<>(currentSet));
                                    currentSet.clear();
                                }
                                currentSet.add(outputRowNum);
                            }

                            Row templateRow = templateSheet.createRow(outputRowNum++);
                            if (!cellSchoolID.equals(previousSchoolID)) {
                                createStyledCell(templateRow, 0, String.valueOf(listNum), templateStyles[0]);
                                listNum++;
                            }

                            previousSchoolID = cellSchoolID;

                            double dblSchoolID = Double.parseDouble(cellSchoolID);
                            int intSchoolID = (int) dblSchoolID;

                            createStyledCell(templateRow, 1, cellDivision, templateStyles[1]);
                            createStyledCell(templateRow, 2, "", templateStyles[2]);
                            createStyledCell(templateRow, 3, String.valueOf(intSchoolID), templateStyles[3]);
                            createStyledCell(templateRow, 4, cellSchoolName, templateStyles[4]);
                            createStyledCell(templateRow, 5, cellGradeLevel, templateStyles[5]);

                            for (int i = 0; i < selectedCheckBoxValues.size(); i++) {
                                String cellValue;
                                if (i < setOfItem.size()) {
                                    cellValue = String.valueOf(setOfItem.get(i));
                                } else {
                                    cellValue = "";
                                }
                                createStyledCell(templateRow, 6 + i, cellValue, templateStyles[6 + i]);
                                System.out.println("setOfItem.get(i): " + cellValue);
                            }


                            createStyledCell(templateRow, 11, "", templateStyles[11]);
                            createStyledCell(templateRow, 12, "", templateStyles[12]);
                            createStyledCell(templateRow, 13, "", templateStyles[13]);
                            createStyledCell(templateRow, 14, "", templateStyles[14]);

                            if (rowCount == maxRowCount) {
                                int rowIndex = 37;
                                templateSheet.shiftRows(rowIndex, templateSheet.getLastRowNum(), 35);
                            }
                            rowCount++;
                        }
                    }
                }

                if (!currentSet.isEmpty()) {
                    sameSchoolIDRows.add(new ArrayList<>(currentSet));
                    currentSet.clear();
                }

                previousSchoolID = null;
            }

            for (List<Integer> rowSet : sameSchoolIDRows) {
                if (rowSet.size() > 1) {
                    for (int colNum = 0; colNum < 5; colNum++) {
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowSet.get(0), rowSet.get(rowSet.size() - 1), colNum, colNum);
                        templateSheet.addMergedRegion(cellRangeAddress);
                    }
                }
            }

            Sheet secondSheet = templateWorkbook.createSheet("Summary");
            int currentRow = 0;

            for (int selectedLot : selectedCheckBoxValues) {
                for (LotSummary_Model model : lotRowsModel) {
                    if (selectedLot == model.getLot()) {
                        int startRow = model.getStartRow();
                        int endRow = model.getEndRow();

                        int copiedStartRow = currentRow;
                        currentRow = copyRows(summaryTemplateSheet, secondSheet, startRow, endRow, currentRow, templateWorkbook);
                        int copiedEndRow = currentRow - 1;

                        List<SummaryItems_Model> summaryItemsModels = new ArrayList<>();
                        getSummaryItems(secondSheet, templateSheet, itemBasesWorkbook, copiedStartRow, copiedEndRow, summaryItemsModels);

                        for (SummaryItems_Model ItemModel : summaryItemsModels) {
                            System.out.println("LOT: " + ItemModel.getLOT());
                            for (GradeLevelSets_Model gradeLevelSetsModel : ItemModel.getGradeLevelSetsModels()) {
                                System.out.println("Grade Level: " + gradeLevelSetsModel.getGradeLevel());
                                for (ItemSets_Model itemSetsModel : gradeLevelSetsModel.getItemSetsModelList()) {
                                    System.out.println("Item Name: " + itemSetsModel.getItemName() + " - " + itemSetsModel.getQty());
                                }
                            }
                        }

                        setData(secondSheet, copiedStartRow, copiedEndRow, summaryItemsModels);
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                templateWorkbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void findHeaderColumnRange(Sheet sheet, Workbook setOfItemsSheet, String targetHeader, Cell cellSchoolID, List<Integer> setOfItem) {
        List<CellRangeAddress> mergedRegions = new ArrayList<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            if (mergedRegion.getFirstRow() == 0) {
                mergedRegions.add(mergedRegion);
            }
        }
        mergedRegions.sort(Comparator.comparingInt(CellRangeAddress::getFirstColumn));

        for (CellRangeAddress cellData : mergedRegions) {
            String title = Optional.ofNullable(sheet.getRow(0).getCell(cellData.getFirstColumn()))
                    .map(Cell::toString)
                    .orElse("Untitled");

            String lot = title.split(":")[0].trim();

            Sheet basesSheet = setOfItemsSheet.getSheet(lot);

            boolean itemQtyFound = false;
            int cellCount = 0;

            if (lot.equals(targetHeader)) {
                // Get the value of the first cell under the merged cell in row 2
                Row row = sheet.getRow(1); // Row 2 (index 1)
                if (row != null) {

                    while (!itemQtyFound) {
                        Cell cell = row.getCell(cellData.getFirstColumn() + cellCount);
                        if (cell != null) {
                            String value = cell.toString();

                            // Get the row index of cellSchoolID
                            int schoolIDRowIndex = cellSchoolID.getRowIndex();

                            // Get the value of the first cell under the merged cell in the row of cellSchoolID
                            Row schoolIDRow = sheet.getRow(schoolIDRowIndex);
                            if (schoolIDRow != null) {
                                Cell schoolIDCell = schoolIDRow.getCell(cellData.getFirstColumn() + cellCount);
                                if (schoolIDCell != null) {
                                    double itemQty = Double.parseDouble(schoolIDCell.toString());

                                    if (itemQty != 0) {
                                        //System.out.println("Value under the merged cell in row 2: " + value);
                                       // System.out.println("Value under the merged cell in row of cellSchoolID: " + itemQty);

                                        // Get the value of column E (Grade Level)
                                        Cell columnECell = schoolIDRow.getCell(4); // Column E is the 4th index (0-based)
                                        if (columnECell != null) {
                                            String columnEValue = columnECell.toString();
                                            //System.out.println("Value of column E in the schoolIDRowIndex: " + columnEValue);

                                            // Search for columnEValue in the basesSheet
                                            boolean valueFound = false;
                                            int endRowIndex = -1; // Initialize endRowIndex

                                            for (int startRowIndex = 0; startRowIndex <= basesSheet.getLastRowNum(); startRowIndex++) {
                                                Row basesRow = basesSheet.getRow(startRowIndex);
                                                if (basesRow != null) {
                                                    Cell cellA = basesRow.getCell(0); // Column A is the 0th index (0-based)
                                                    if (cellA != null && columnEValue.equals(cellA.toString())) {
                                                        //System.out.println("Found columnEValue in basesSheet at row: " + startRowIndex);
                                                        valueFound = true;

                                                        // Determine endRowIndex
                                                        int tempRowIndex = startRowIndex;
                                                        boolean isEmpty = false;
                                                        while (!isEmpty && tempRowIndex <= basesSheet.getLastRowNum()) {
                                                            Row findEmptyRow = basesSheet.getRow(tempRowIndex);
                                                            if (findEmptyRow == null || findEmptyRow.getPhysicalNumberOfCells() == 0) {
                                                                endRowIndex = tempRowIndex - 1; // Last non-empty row
                                                                //System.out.println("End at row: " + endRowIndex);
                                                                isEmpty = true;
                                                            } else {
                                                                boolean rowIsEmpty = true;
                                                                for (Cell cell1 : findEmptyRow) {
                                                                    if (cell1.getCellType() != CellType.BLANK) {
                                                                        rowIsEmpty = false;
                                                                        break;
                                                                    }
                                                                }
                                                                if (rowIsEmpty) {
                                                                    endRowIndex = tempRowIndex - 1; // Last non-empty row
                                                                    //System.out.println("End at row: " + endRowIndex);
                                                                    isEmpty = true;
                                                                }
                                                            }
                                                            tempRowIndex++;
                                                        }

                                                        // Search within the range from startRowIndex to endRowIndex
                                                        if (endRowIndex != -1) {
                                                            for (int searchRowIndex = startRowIndex; searchRowIndex <= endRowIndex; searchRowIndex++) {
                                                                Row searchRow = basesSheet.getRow(searchRowIndex);
                                                                if (searchRow != null) {
                                                                    Cell searchCellA = searchRow.getCell(0);
                                                                    if (searchCellA != null && value.equals(searchCellA.toString())) {
                                                                        //System.out.println("Found value in the range at row: " + searchRowIndex);
                                                                        // Perform further actions as needed
                                                                        // Retrieve value from Cell B (Column 1)
                                                                        Cell cellB = searchRow.getCell(1); // Column B is the 1st index (0-based)
                                                                        if (cellB != null) {
                                                                            double setOfItemValue = Double.parseDouble(cellB.toString());
                                                                            //System.out.println("Value of Cell B in row " + searchRowIndex + ": " + setOfItemValue);

                                                                            // Perform further actions with cellBValue if needed

                                                                            int set = (int) (itemQty / setOfItemValue);
                                                                            setOfItem.add(set); // Add the value to the list
                                                                        } else {
                                                                            setOfItem.add(0);
                                                                            //System.out.println("Cell B (Column 1) is null in row: " + searchRowIndex);
                                                                        }

                                                                        // Perform further actions if needed and break out of loop
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        break; // Exit loop once value is found
                                                    }
                                                }
                                            }
                                            if (!valueFound) {
                                                System.out.println("Value of columnEValue not found in basesSheet.");
                                            }

                                            itemQtyFound = true;
                                        } else {
                                            System.out.println("Column E in the schoolIDRowIndex is null.");
                                        }

                                    } else {
                                        cellCount++;
                                    }
                                } else {
                                    System.out.println("Cell under the merged cell in row of cellSchoolID is null.");
                                }
                            } else {
                                System.out.println("Row for cellSchoolID does not exist.");
                            }
                        } else {
                           // System.out.println("Cell under the merged cell in row 2 is null.");
                        }
                    }

                } else {
                    System.out.println("Row 2 does not exist.");
                }
            }
        }
    }

    // Method to copy rows from summaryTemplateSheet to secondSheet
    private int copyRows(Sheet sourceSheet, Sheet destinationSheet, int startRow, int endRow, int destinationStartRow, Workbook destinationWorkbook) {
        // Copy column widths
        for (int colNum = 0; colNum < sourceSheet.getRow(0).getLastCellNum(); colNum++) {
            destinationSheet.setColumnWidth(colNum, sourceSheet.getColumnWidth(colNum));
        }

        // Copy rows
        for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
            Row sourceRow = sourceSheet.getRow(rowNum);
            if (sourceRow != null) {
                Row destinationRow = destinationSheet.createRow(destinationStartRow + (rowNum - startRow));
                for (int colNum = 0; colNum < sourceRow.getLastCellNum(); colNum++) {
                    Cell sourceCell = sourceRow.getCell(colNum);
                    if (sourceCell != null) {
                        Cell destinationCell = destinationRow.createCell(colNum, sourceCell.getCellType());
                        switch (sourceCell.getCellType()) {
                            case STRING:
                                destinationCell.setCellValue(sourceCell.getStringCellValue());
                                break;
                            case NUMERIC:
                                destinationCell.setCellValue(sourceCell.getNumericCellValue());
                                break;
                            case BOOLEAN:
                                destinationCell.setCellValue(sourceCell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                destinationCell.setCellFormula(sourceCell.getCellFormula());
                                break;
                            case ERROR:
                                destinationCell.setCellErrorValue(sourceCell.getErrorCellValue());
                                break;
                            default:
                                break;
                        }
                        // Create a new cell style in the destination workbook
                        CellStyle newStyle = destinationWorkbook.createCellStyle();
                        newStyle.cloneStyleFrom(sourceCell.getCellStyle()); // Clone the style
                        destinationCell.setCellStyle(newStyle); // Apply the new style
                    }
                }
            }
        }

        return destinationStartRow + (endRow - startRow + 3); // Return the next available row index for the next copy operation
    }

    public void getSummaryItems(Sheet secondSheet, Sheet templateSheet, Workbook itemBasesWorkbook, int startRow, int endRow, List<SummaryItems_Model> models) {
        Row row = secondSheet.getRow(startRow);
        if (row == null) return;

        Cell firstCell = row.getCell(0);
        if (firstCell != null) {
            String lotValue = getCellValue(firstCell);
            Sheet basesSheet = null;
            List<GradeLevelSets_Model> gradeLevelSetsModels = new ArrayList<>();

            // Find the column index where lotValue matches values in G11 to K11
            int lotValueColumnIndex = -1;
            Row templateHeaderRow = templateSheet.getRow(10); // Row 11 (0-based index 10)
            for (int colIndex = 6; colIndex <= 10; colIndex++) { // Columns G to K
                Cell headerCell = templateHeaderRow.getCell(colIndex);
                if (headerCell != null) {
                    String headerValue = getCellValue(headerCell);
                    if (lotValue.equals(headerValue)) {
                        lotValueColumnIndex = colIndex;
                        basesSheet = itemBasesWorkbook.getSheet(lotValue);
                        break;
                    }
                }
            }

            if (lotValueColumnIndex == -1 || basesSheet == null) {
                return; // Exit method if no match found or basesSheet is null
            }

            // Loop through columns 2 to 6 (index 1 to 5) for grade levels
            for (int colIndex = 1; colIndex <= 5; colIndex++) {
                Cell gradeLevelCell = row.getCell(colIndex);
                if (gradeLevelCell != null) {
                    String gradeLevelValue = getCellValue(gradeLevelCell);
                    List<ItemSets_Model> itemSetsModels = new ArrayList<>();
                    int totalItems = 0;

                    // Search for gradeLevelValue in templateSheet from F13 downwards
                    int gradeLevelRowIndex = 12; // Row 13
                    while (true) {
                        Row gradeLevelRow = templateSheet.getRow(gradeLevelRowIndex);
                        if (gradeLevelRow == null) break;

                        Cell gradeLevelCellInTemplate = gradeLevelRow.getCell(5); // Column F
                        if (gradeLevelCellInTemplate != null) {
                            String gradeLevelValueInTemplate = getCellValue(gradeLevelCellInTemplate);
                            if (gradeLevelValue.equals(gradeLevelValueInTemplate)) {
                                Cell intersectingCell = gradeLevelRow.getCell(lotValueColumnIndex);
                                if (intersectingCell != null) {
                                    try {
                                        if (!getCellValue(intersectingCell).isEmpty()) {
                                            int intersectingValue = Integer.parseInt(getCellValue(intersectingCell));
                                            totalItems += intersectingValue;
                                        }
                                    } catch (NumberFormatException e) {
                                        // Handle or log number format exceptions
                                        e.printStackTrace();
                                    }
                                }
                            }
                        }
                        gradeLevelRowIndex++;
                    }

                    System.out.println("lotValue: " + lotValue + " - gradeLevelValue: " + gradeLevelValue + " - totalItems: " + totalItems);

                    // Loop through rows starting from the second row to get item names and sets
                    for (int itemRowIndex = startRow + 1; itemRowIndex <= endRow; itemRowIndex++) {
                        Row itemRow = secondSheet.getRow(itemRowIndex);
                        if (itemRow != null) {
                            Cell itemNameCell = itemRow.getCell(0); // Item names in column 1
                            if (itemNameCell != null) {
                                String itemName = getCellValue(itemNameCell).replace("• ", "");
                                boolean valueFound = false;
                                int endRowIndex = -1;
                                double totalSet = 0; // Changed to double to handle decimal values

                                if (itemName.equals("Motherbox")){
                                    totalSet = totalItems;
                                }
                                else {
                                    // Search for itemName in basesSheet
                                    for (int startRowIndex = 0; startRowIndex <= basesSheet.getLastRowNum(); startRowIndex++) {
                                        Row basesRow = basesSheet.getRow(startRowIndex);
                                        if (basesRow != null) {
                                            Cell cellA = basesRow.getCell(0); // Column A
                                            if (cellA != null && gradeLevelValue.equals(getCellValue(cellA))) {
                                                valueFound = true;

                                                // Determine endRowIndex
                                                int tempRowIndex = startRowIndex;
                                                boolean isEmpty = false;
                                                while (!isEmpty && tempRowIndex <= basesSheet.getLastRowNum()) {
                                                    Row findEmptyRow = basesSheet.getRow(tempRowIndex);
                                                    if (findEmptyRow == null || findEmptyRow.getPhysicalNumberOfCells() == 0) {
                                                        endRowIndex = tempRowIndex - 1;
                                                        isEmpty = true;
                                                    } else {
                                                        boolean rowIsEmpty = true;
                                                        for (Cell cell1 : findEmptyRow) {
                                                            if (cell1.getCellType() != CellType.BLANK) {
                                                                rowIsEmpty = false;
                                                                break;
                                                            }
                                                        }
                                                        if (rowIsEmpty) {
                                                            endRowIndex = tempRowIndex - 1;
                                                            isEmpty = true;
                                                        }
                                                    }
                                                    tempRowIndex++;
                                                }

                                                // Search within the range
                                                if (endRowIndex != -1) {
                                                    for (int searchRowIndex = startRowIndex; searchRowIndex <= endRowIndex; searchRowIndex++) {
                                                        Row searchRow = basesSheet.getRow(searchRowIndex);
                                                        if (searchRow != null) {
                                                            Cell searchCellA = searchRow.getCell(0);
                                                            if (searchCellA != null && itemName.equals(getCellValue(searchCellA).replace("• ", ""))) {
                                                                Cell cellB = searchRow.getCell(1); // Column B
                                                                if (cellB != null) {
                                                                    try {
                                                                        double setOfItemValue = Double.parseDouble(getCellValue(cellB));
                                                                        totalSet = totalItems * setOfItemValue;
                                                                        System.out.println("ItemName: " + itemName + " - QTY: " + setOfItemValue + " - TotalSet: " + totalSet);
                                                                    } catch (NumberFormatException e) {
                                                                        // Handle or log number format exceptions
                                                                        e.printStackTrace();
                                                                    }
                                                                }
                                                                break;
                                                            }
                                                        }
                                                    }
                                                }
                                                break;
                                            }
                                        }
                                    }
                                }
                                itemSetsModels.add(new ItemSets_Model(itemName.replace("• ", ""), (int) totalSet)); // Casting to int if needed
                            }
                        }
                    }
                    gradeLevelSetsModels.add(new GradeLevelSets_Model(gradeLevelValue, itemSetsModels));
                }
            }
            models.add(new SummaryItems_Model(lotValue, gradeLevelSetsModels));
        }
    }


    public void setData(Sheet secondSheet, int startRow, int endRow, List<SummaryItems_Model> summaryItemsModels) {
        for (SummaryItems_Model summaryItemsModel : summaryItemsModels) {
            String lot = summaryItemsModel.getLOT();
            List<GradeLevelSets_Model> gradeLevelSetsModels = summaryItemsModel.getGradeLevelSetsModels();

            for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
                Row row = secondSheet.getRow(rowIndex);
                if (row == null) continue;

                Cell itemCell = row.getCell(0);
                if (itemCell != null) {
                    String itemName = getCellValue(itemCell).replace("• ", "");

                    for (GradeLevelSets_Model gradeLevelSetsModel : gradeLevelSetsModels) {
                        String gradeLevel = gradeLevelSetsModel.getGradeLevel();
                        List<ItemSets_Model> itemSetsModels = gradeLevelSetsModel.getItemSetsModelList();

                        for (ItemSets_Model itemSetsModel : itemSetsModels) {
                            if (itemName.equals(itemSetsModel.getItemName().replace("• ", ""))) {
                                int qty = itemSetsModel.getQty();

                                // Determine the correct column based on grade level
                                int gradeLevelColumnIndex = getGradeLevelColumnIndex(gradeLevel);

                                if (gradeLevelColumnIndex != -1) {
                                    Cell qtyCell = row.getCell(gradeLevelColumnIndex);
                                    if (qtyCell == null) {
                                        qtyCell = row.createCell(gradeLevelColumnIndex);
                                    }

                                    if (qty != 0) {
                                        qtyCell.setCellValue(qty);
                                    } else {
                                        qtyCell.setCellValue("");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        // Additional process to find rows with the same item name and sum values
        for (int rowIndex = startRow+1; rowIndex <= endRow; rowIndex++) {
            Row row = secondSheet.getRow(rowIndex);
            if (row == null) continue;

            Cell itemCell = row.getCell(0);
            if (itemCell != null) {
                String itemName = getCellValue(itemCell).replace("• ", "").trim();

                double sum = 0;
                for (int innerRowIndex = startRow; innerRowIndex <= endRow; innerRowIndex++) {
                    Row innerRow = secondSheet.getRow(innerRowIndex);
                    if (innerRow == null) continue;

                    Cell innerItemCell = innerRow.getCell(0);
                    if (innerItemCell != null && itemName.equals(getCellValue(innerItemCell).replace("• ", "").trim())) {
                        for (int colIndex = 1; colIndex <= 5; colIndex++) { // Columns B to F (1-based)
                            Cell valueCell = innerRow.getCell(colIndex);
                            if (valueCell != null && valueCell.getCellType() == CellType.NUMERIC) {
                                sum += valueCell.getNumericCellValue();
                            }
                        }
                    }
                }

                // Set the sum to column G (6, 0-based)
                Cell sumCell = row.getCell(6);
                if (sumCell == null) {
                    sumCell = row.createCell(6);
                }
                sumCell.setCellValue(sum);
            }
        }
    }

    // Helper method to get the column index based on grade level
    private int getGradeLevelColumnIndex(String gradeLevel) {
        return switch (gradeLevel) {
            case "G1toG3" -> 1; // Column B
            case "G4toG6" -> 2; // Column C
            case "JHS" -> 3; // Column D
            case "SHSCore" -> 4; // Column E
            case "SHSStem" -> 5; // Column F
            default -> -1; // Invalid grade level
        };
    }


    // Method to get cell value as a string
    private static String getCellValue(Cell cell) {
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue().toString();
                    } else {
                        return String.valueOf(cell.getNumericCellValue());
                    }
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    return cell.getCellFormula();
                case BLANK:
                    return "";
                default:
                    return "Unsupported cell type";
            }
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

