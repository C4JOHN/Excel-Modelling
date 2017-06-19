
package complexexcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Iterator;
import javafx.application.Application;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.scene.Scene;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.Tooltip;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.BorderPane;
import javafx.scene.paint.Color;
import javafx.stage.Stage;
import javafx.util.Callback;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

public class ComplexExcelDisplay extends Application {
   
    public static ObservableList<CellModel> cells =FXCollections.observableArrayList();
    public static int x=0;
    public static int y=0;
    @Override
    public void start(Stage primaryStage) throws IOException {
        
    SimpleDateFormat sdf=new SimpleDateFormat("dd/MM/yyy");
        HSSFWorkbook workbook=new HSSFWorkbook(new FileInputStream(
                "C:\\Users\\removevirus\\Desktop\\complex1.xls"));
Iterator<org.apache.poi.ss.usermodel.Sheet> sheetIterator=workbook.iterator();
        
        while(sheetIterator.hasNext()){
            HSSFSheet sheet=(HSSFSheet) sheetIterator.next();
            Iterator<Row> rowIterator=sheet.iterator();
            while(rowIterator.hasNext()){
               
                Row row=rowIterator.next();
                Iterator<org.apache.poi.ss.usermodel.Cell> cellIterator=row.iterator();
                while(cellIterator.hasNext()){
                    HSSFCell cell=(HSSFCell) cellIterator.next();
                    cells.add(new CellModel(cell.getAddress(),cell.getCellTypeEnum(),cell));
                }
            }
        }
    
    TableView<CellModel> view=new TableView();
    view.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
    view.setItems(cells);
        
    TableColumn<CellModel,HSSFCell> stringColumn=new TableColumn<>("Cells");
    stringColumn.setCellValueFactory(new PropertyValueFactory<>("cell"));
    
    
    stringColumn.setCellFactory(new Callback<TableColumn<CellModel,HSSFCell>,TableCell<CellModel,HSSFCell>>(){
   TableCell<CellModel,HSSFCell> cell=new TableCell();
  
        @Override
        public TableCell<CellModel, HSSFCell> call(TableColumn<CellModel, HSSFCell> param) {
            return new TableCell<CellModel,HSSFCell>(){
                @Override
                protected void updateItem(HSSFCell item,boolean empty){
                
                    if(!empty){
                        setText(item.toString());
                        Tooltip tip=new Tooltip();
                        setTooltip(tip);
                        tip.setText(item.getAddress().toString());
                    }
                    
                    else{
                       setText(null);
                    }
                }
            };
        }
        
    });
    
    System.out.println(cells.size());
    
    view.getColumns().add(stringColumn);
    
    BorderPane pane=new BorderPane();
    Scene scene=new Scene(pane,1000,500,Color.rgb(0, 0, 0, 0));
   // scene.getStylesheets().add(getClass().getResource("ExcellCss.css").toExternalForm());
    pane.setCenter(view);
    primaryStage.setScene(scene);
    primaryStage.show();
    }
            
    public static void main(String[] args) {
        launch(args);
        
    }
}
