package com.teclan.poi.word;

public class Person extends org.javalite.activejdbc.Model{

    private String name;
    private String sex;
    private String type;
    private String idCard;
    private String location;
    private String enterTime;
    private String outTime;
    private String contNo;
    private String wage;
    private String cardNo;
    private String bank;
    private String phone;
    private String memo;

    public Person() {
    }

    public Person(String name, String sex, String type, String idCard, String location, String enterTime, String outTime, String contNo, String wage, String cardNo, String bank, String phone, String memo) {
        this.name = name;
        this.sex = sex;
        this.type = type;
        this.idCard = idCard;
        this.location = location;
        this.enterTime = enterTime;
        this.outTime = outTime;
        this.contNo = contNo;
        this.wage = wage;
        this.cardNo = cardNo;
        this.bank = bank;
        this.phone = phone;
        this.memo = memo;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getIdCard() {
        return idCard;
    }

    public void setIdCard(String idCard) {
        this.idCard = idCard;
    }

    public String getLocation() {
        return location;
    }

    public void setLocation(String location) {
        this.location = location;
    }

    public String getEnterTime() {
        return enterTime;
    }

    public void setEnterTime(String enterTime) {
        this.enterTime = enterTime;
    }

    public String getOutTime() {
        return outTime;
    }

    public void setOutTime(String outTime) {
        this.outTime = outTime;
    }

    public String getContNo() {
        return contNo;
    }

    public void setContNo(String contNo) {
        this.contNo = contNo;
    }

    public String getWage() {
        return wage;
    }

    public void setWage(String wage) {
        this.wage = wage;
    }

    public String getCardNo() {
        return cardNo;
    }

    public void setCardNo(String cardNo) {
        this.cardNo = cardNo;
    }

    public String getBank() {
        return bank;
    }

    public void setBank(String bank) {
        this.bank = bank;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public String getMemo() {
        return memo;
    }

    public void setMemo(String memo) {
        this.memo = memo;
    }
}
