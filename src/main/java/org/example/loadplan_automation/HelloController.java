package org.example.loadplan_automation;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.ResourceBundle;
import java.util.Set;
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

            readExcelFile();
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
                String urlString = "http://192.168.1.229/NIKKA/get_data.php";
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

                // Print the response content for debugging
                System.out.println(content);

                // Parse JSON data
                JSONObject json = new JSONObject(content.toString());

                for (String tableName : json.keySet()) {
                    JSONArray tableData = json.getJSONArray(tableName);
                    for (int i = 0; i < tableData.length(); i++) {
                        JSONObject row = tableData.getJSONObject(i);
                        Region_model regionModel = new Region_model(
                                row.getString("Region"),
                                row.getString("Division"),
                                row.getInt("SchoolID"),
                                row.getString("School Name"),
                                row.getString("GradeLevel")
                        );
                        list.add(regionModel);

                        regions.add(row.getString("Region"));
                        allDivisions.add(row.getString("Division"));
                        allGradeLevels.add(row.getString("GradeLevel"));
                    }
                }

                // Update ComboBoxes and TableView on JavaFX Application Thread
                javafx.application.Platform.runLater(() -> {
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

    private void readExcelFile() {
        String filePath = "C:\\Users\\5CG6105SVT\\Desktop\\LTE-SM-2023-Allocation-List-Lots-1415161718-Nikka-Trading.xlsx";

        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = WorkbookFactory.create(fis)) {

            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate through each row
            for (Row row : sheet) {
                // Iterate through each cell in the row
                for (Cell cell : row) {
                    // Print the cell value
                    System.out.print(cell.toString() + "\t");
                }
                System.out.println(); // Move to the next line for the next row
            }
        } catch (IOException | EncryptedDocumentException ex) {
            ex.printStackTrace();
        }
    }

}
