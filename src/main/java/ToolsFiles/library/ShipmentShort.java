package ToolsFiles.library;

import java.text.SimpleDateFormat;
import java.util.Date;

public class ShipmentShort implements Comparable {

    private long inn;
    private Date shipDate;
    private double summ;
    private String name;

    public ShipmentShort(Shipment shipment) {
        this.inn = shipment.getInn();
        this.shipDate = shipment.getShipDate();
        this.summ = shipment.getSumm();
    }

    public ShipmentShort(Shipment shipment, String name) {
        this.inn = shipment.getInn();
        this.shipDate = shipment.getShipDate();
        this.summ = shipment.getSumm();
        this.name = name;
    }

    public ShipmentShort(Long inn, Date date, double summ) {
        this.inn = inn;
        this.shipDate = date;
        this.summ = summ;
    }

    public String getCustomer() {
        if (inn == 7706092528l) {
            return Test.getRightNameForCashClient(name, Test.createCommonNameForCashClientMap());
        } else {
            return Test.getNameOfClientByInn(inn);
        }
    }

    public Date getShipDate() {
        return shipDate;
    }

    public long getInn() {
        return inn;
    }

    public double getSumm() {
        return summ;
    }

    public String createSimpleDate(Date date) {

        SimpleDateFormat format = new SimpleDateFormat("dd.MM.yyyy");
        String normalDate = format.format(date);
        return normalDate;
    }

    @Override
    public String toString() {

        return "Отгрузка " + createSimpleDate(getShipDate()) + ", " + getCustomer() + ", " + summ;
    }

    @Override
    public int compareTo(Object o) {
        int result;

        result = this.getShipDate().compareTo(((ShipmentShort) o).getShipDate());

        if (result == 0) {
            result = 1;
        }

        return result;
    }
}
