package ToolsFiles.library;

import java.util.Date;
import java.util.HashMap;

public class Shipment implements Comparable{

    private String customer, data;
    private Date shipDate;
    private long inn;
    private double summ;
    //private ArrayList<ArrayList> order;
    private HashMap<String, Integer> order;
    private boolean isCash;

    public Shipment(String customer, Date shipDate, HashMap<String, Integer> order, long inn, boolean isCash, double summ) {
        this.customer = customer;
        this.shipDate = shipDate;
        this.order = order;
        this.inn = inn;
        this.isCash = isCash;
        this.summ = summ;
    }

    public String getCustomer() {
        return customer;
    }
    public Date getShipDate() {
        return shipDate;
    }
    public HashMap<String, Integer> getOrder() {return order;}
    public boolean getIsCash() {
        return isCash;
    }
    public long getInn() {
        return inn;
    }
    public double getSumm() {
        return summ;
    }


    @Override
    public int compareTo(Object o) {
        int result;

        result =  this.getShipDate().compareTo(((Shipment) o).getShipDate());

        if (result == 0) {

            result =  this.getCustomer().compareTo(((Shipment) o).getCustomer());
        }

        if (result == 0) {

            result =  1;
        }
        return result;
    }
}
