package org.traccar.reports;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.traccar.model.Device;
import org.traccar.model.User;
import org.traccar.reports.common.ReportUtils;
import org.traccar.reports.model.UserDeviceItem;
import org.traccar.storage.Storage;
import org.traccar.storage.StorageException;
import org.traccar.storage.query.Columns;
import org.traccar.storage.query.Condition;
import org.traccar.storage.query.Request;

import javax.inject.Inject;
import java.io.*;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

public class DevicesReportProvider {

    private final Storage storage;
    private final String EMAIL = "EMAIL";
    private final String DEVICE_NAME = "UREĐAJ";
    private final String TOTAL = "TOTAL";

    private final ReportUtils reportUtils;

    @Inject
    public DevicesReportProvider(Storage storage, ReportUtils reportUtils) {
        this.storage = storage;
        this.reportUtils = reportUtils;
    }

    public void getDevicesInfo(OutputStream stream) throws StorageException, IOException {
        List<User> users = storage.getObjects(User.class, new Request(new Columns.All()))
                .stream()
                .filter(user -> !user.getAdministrator())
                .collect(Collectors.toList());
        List<UserDeviceItem> usersWithDevices = new ArrayList<>();
        Map<Long, Device> countedDevices = new HashMap<>();

        for(User user: users) {
            var conditions = new LinkedList<Condition>();
            conditions.add(new Condition.Permission(User.class, user.getId(), Device.class).excludeGroups());
            List<Device> devices = new ArrayList<>(storage.getObjects(Device.class, new Request(new Columns.All(), Condition.merge(conditions))));
            countDevices(devices, countedDevices);
            usersWithDevices.add(new UserDeviceItem(user, devices));
        }

        createExcel(usersWithDevices, countedDevices).write(stream);
    }

    private XSSFWorkbook createExcel(List<UserDeviceItem> usersWithDevices, Map<Long, Device> countedDevices) {
        XSSFWorkbook workbook = new XSSFWorkbook();

        List<Long> alreadyWrittenDevices = new ArrayList<>();

        XSSFSheet spreadsheet = workbook.createSheet(" UREĐAJI ");
        spreadsheet.setColumnWidth(0, 10000);
        spreadsheet.setColumnWidth(1, 10000);
        XSSFRow row;
        createHeader(spreadsheet);
        int rowId = 1;
        for(UserDeviceItem userDeviceItem : usersWithDevices) {
            row = createRow(spreadsheet, rowId++);
            int col = 0;
            XSSFCell cell = row.createCell(col++);
            cell.setCellValue(userDeviceItem.getUser().getEmail());
            applyCellStyle(workbook, cell, EMAIL, true, true, false);
            if(userDeviceItem.getDevices().isEmpty()) continue;
            cell = row.createCell(col);
            Device firstDevice = userDeviceItem.getDevices().get(0);
            cell.setCellValue(firstDevice.getName());
            applyCellStyle(workbook, cell, DEVICE_NAME, true, checkIndex(null, userDeviceItem.getDevices()), alreadyWrittenDevices.contains(firstDevice.getId()));
            alreadyWrittenDevices.add(firstDevice.getId());
            for(Device device: userDeviceItem.getDevices().subList(1, userDeviceItem.getDevices().size())) {
                row = createRow(spreadsheet, rowId++);
                row.createCell(0);
                cell = row.createCell(col);
                cell.setCellValue(device.getName());
                applyCellStyle(workbook, cell, DEVICE_NAME, false, checkIndex(device, userDeviceItem.getDevices()), alreadyWrittenDevices.contains(device.getId()));
                alreadyWrittenDevices.add(device.getId());
            }
        }
        createTotal(spreadsheet, countedDevices);
        return workbook;
    }

    private boolean checkIndex(Device device, List<Device> devices) {
        if(device == null) {
            return devices.size() == 1;
        }
        return devices.indexOf(device) == devices.size() - 1;
    }

    private void applyCellStyle(Workbook workbook, XSSFCell cell, String type, boolean top, boolean bottom, boolean alreadyWritten) {
        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
        if(EMAIL.equals(type)) {
            setBorderStyle(style, top, bottom, IndexedColors.BLUE.index);
            style.setFillPattern(FillPatternType.SQUARES);
            style.setFillForegroundColor(IndexedColors.SKY_BLUE.index);
            applyFont((XSSFWorkbook) workbook, style, (short) 12);
        }
        if(DEVICE_NAME.equals(type)) {
            if(alreadyWritten) {
                setBorderStyle(style, top, bottom, IndexedColors.GREY_80_PERCENT.index);
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                applyFont((XSSFWorkbook) workbook, style, (short) 12);
            } else {
                setBorderStyle(style, top, bottom, IndexedColors.GREEN.index);
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
                applyFont((XSSFWorkbook) workbook, style, (short) 12);
            }

        }
        if(TOTAL.equals(type)) {
            setBorderStyle(style, top, bottom, IndexedColors.GREEN.index);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
            applyFont((XSSFWorkbook) workbook, style, (short) 12);
        }
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(style);
    }

    private void applyFont(XSSFWorkbook workbook, XSSFCellStyle style, short height) {
        XSSFFont font= workbook.createFont();
        font.setFontHeightInPoints(height);
        font.setFontName("Arial");
        font.setColor(IndexedColors.BLACK.getIndex());
        font.setBold(true);
        font.setItalic(false);

        style.setFont(font);
    }

    private void setBorderStyle(XSSFCellStyle style, boolean top, boolean bottom, short color) {
        if(top) {
            style.setBorderTop(BorderStyle.THICK);
            style.setTopBorderColor(color);
        }
        if(bottom) {
            style.setBorderBottom(BorderStyle.THICK);
            style.setBottomBorderColor(color);
        } else {
            style.setBorderBottom(BorderStyle.DASHED);
            style.setBottomBorderColor(color);
        }
        style.setBorderLeft(BorderStyle.THICK);
        style.setLeftBorderColor(color);

        style.setBorderRight(BorderStyle.THICK);
        style.setRightBorderColor(color);
    }


    private void createTotal(XSSFSheet spreadsheet, Map<Long, Device> countedDevices) {
        XSSFRow row = spreadsheet.getRow(2);
        int col;
        for(col = 2; col <= 7; col++) {
            row.createCell(col);
        }
        Cell cell = row.createCell(col++);
        cell.setCellValue("Total");
        applyCellStyle(spreadsheet.getWorkbook(), (XSSFCell) cell, TOTAL, true, true, false);


        cell = row.createCell(col++);
        cell.setCellValue(countedDevices.size());
        applyCellStyle(spreadsheet.getWorkbook(), (XSSFCell) cell, TOTAL, true, true, false);

        cell = row.createCell(col);
        cell.setCellValue(countedDevices.size() * 60 + " RSD");
        applyCellStyle(spreadsheet.getWorkbook(), (XSSFCell) cell, TOTAL, true, true, false);

    }

    private void createHeader(XSSFSheet spreadsheet) {
        XSSFRow row = createRow(spreadsheet, 0);
        int col = 0;
        String[] headers = {EMAIL, DEVICE_NAME};
        for(String header : headers) {
            XSSFCell cell = row.createCell(col++);
            cell.setCellValue(header);
            applyCellStyle(spreadsheet.getWorkbook(), cell, header, true, false, false);
            applyFont(spreadsheet.getWorkbook(), cell.getCellStyle(), (short) 14);
        }
    }

    private void countDevices(List<Device> devices, Map<Long, Device> countedDevices) {
        devices.forEach(device -> countedDevices.put(device.getId(), device));
    }

    private XSSFRow createRow(XSSFSheet spreadsheet, int rowNum) {
        XSSFRow row = spreadsheet.createRow(rowNum);
        row.setHeight((short) 500);
        return row;
    }
}
