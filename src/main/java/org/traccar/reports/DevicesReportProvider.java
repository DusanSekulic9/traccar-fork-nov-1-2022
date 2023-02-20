package org.traccar.reports;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.traccar.api.resource.DeviceResource;
import org.traccar.model.Device;
import org.traccar.model.User;
import org.traccar.reports.model.UserDeviceItem;
import org.traccar.storage.Storage;
import org.traccar.storage.StorageException;
import org.traccar.storage.query.Columns;
import org.traccar.storage.query.Condition;
import org.traccar.storage.query.Request;

import javax.inject.Inject;
import javax.ws.rs.core.HttpHeaders;
import javax.ws.rs.core.Response;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;
import java.util.stream.Collectors;

public class DevicesReportProvider {

    private final Storage storage;
    private final DeviceResource deviceResource;

    @Inject
    public DevicesReportProvider(Storage storage, DeviceResource deviceResource) {
        this.storage = storage;
        this.deviceResource = deviceResource;
    }

    public Response getDevicesInfo() throws StorageException, IOException {
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
            System.out.println(user.getId() + " devices: " + devices.size());
            usersWithDevices.add(new UserDeviceItem(user, devices));
        }

        OutputStream ou = new FileOutputStream("report.xlsx");
        createExcel(usersWithDevices, countedDevices).write(ou);
        System.out.println(countedDevices.size());
        return Response.ok(ou).header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=report.xlsx").build();
    }

    private XSSFWorkbook createExcel(List<UserDeviceItem> usersWithDevices, Map<Long, Device> countedDevices) {
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet spreadsheet = workbook.createSheet(" Devices ");
        XSSFRow row;
        createHeader(spreadsheet);
        int rowId = 1;
        for(UserDeviceItem userDeviceItem : usersWithDevices) {
            row = spreadsheet.createRow(rowId++);
            int col = 0;
            Cell cell = row.createCell(col++);
            cell.setCellValue(userDeviceItem.getUser().getEmail());
            cell = row.createCell(col);
            cell.setCellValue(userDeviceItem.getDevices().get(0).getName());
            for(Device device: userDeviceItem.getDevices().subList(1, userDeviceItem.getDevices().size() - 1)) {
                row = spreadsheet.createRow(rowId++);
                cell = row.createCell(col);
                cell.setCellValue(device.getName());
            }
        }
        createTotal(spreadsheet, countedDevices);
        return workbook;
    }

    private void createTotal(XSSFSheet spreadsheet, Map<Long, Device> countedDevices) {
        XSSFRow row = spreadsheet.getRow(spreadsheet.getLastRowNum() % 7);
        int col = 7;
        Cell cell = row.createCell(col++);
        cell.setCellValue("Total");

        cell = row.createCell(col);
        cell.setCellValue(countedDevices.size());
    }

    private void createHeader(XSSFSheet spreadsheet) {
        XSSFRow row = spreadsheet.createRow(0);
        int col = 0;
        String[] headers = {"Ime", "Uredjaji"};
        for(String header : headers) {
            Cell cell = row.createCell(col++);
            cell.setCellValue(header);
        }
    }

    private void countDevices(List<Device> devices, Map<Long, Device> countedDevices) {
        devices.forEach(device -> countedDevices.put(device.getId(), device));
    }
}
