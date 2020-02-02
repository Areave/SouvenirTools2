package ToolsFiles.library;

public class Client {

    String name;
    long inn;
    String contact;

    public Client(String name, long inn, String contact) {
        this.name = name;
        this.inn = inn;
        this.contact = contact;
    }

    @Override
    public String toString() {
        return "Client{" +
                "name='" + name + '\'' +
                ", inn=" + inn +
                ",\ncontact='" + contact + '\'' +
                "}\n";
    }
}
