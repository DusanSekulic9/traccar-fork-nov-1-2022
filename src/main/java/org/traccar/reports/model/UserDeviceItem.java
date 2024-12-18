package org.traccar.reports.model;

import org.traccar.model.Device;
import org.traccar.model.User;

import java.util.List;

public class UserDeviceItem {

    private final User user;
    private final List<Device> devices;

    public UserDeviceItem(User user, List<Device> devices) {
        this.user = user;
        this.devices = devices;
    }

    public List<Device> getDevices() {
        return devices;
    }

    public User getUser() {
        return user;
    }
}
