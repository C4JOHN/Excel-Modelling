package complexexcel;

import javafx.scene.control.Cell;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellAddress;
public class CellModel extends Cell{
   public CellAddress address;
   public CellType cellType;
   public HSSFCell cell;

    public CellModel(CellAddress address, CellType cellType, HSSFCell cell) {
        this.address = address;
        this.cellType = cellType;
        this.cell = cell;
    }

    public CellAddress getAddress() {
        return address;
    }

    public void setAddress(CellAddress address) {
        this.address = address;
    }

    public CellType getCellType() {
        return cellType;
    }

    public void setCellType(CellType cellType) {
        this.cellType = cellType;
    }

    public HSSFCell getCell() {
        return cell;
    }

    public void setCell(HSSFCell cell) {
        this.cell = cell;
    }
 
   
    
    
}
