package myProject.excelUtil.pojo;

import java.io.Serializable;
import java.math.BigDecimal;
import java.util.Date;

import myProject.excelUtil.annotation.ExcelBigDecimal;
import myProject.excelUtil.annotation.ExcelBoolean;
import myProject.excelUtil.annotation.ExcelDate;
import myProject.excelUtil.annotation.ExcelField;
import myProject.excelUtil.annotation.ExcelValid;

public class OrderForm implements Serializable {

	private static final long serialVersionUID = 2483106904523249992L;
	@ExcelField(name = "订单号", sort = 0, nullable = false)
	@ExcelValid(regexp = "\\d{10}$")
	private String orderNo;

	@ExcelField(name = "姓名", sort = 1, nullable = false)
	private String customerName;

	@ExcelField(name = "支付金额", sort = 2, nullable = false)
	@ExcelValid(regexp = "^[+-]?[0-9]+(.[0-9]{1,2})?$")
	private BigDecimal orderAmt;

	@ExcelField(name = "优惠金额", sort = 3, nullable = false)
	@ExcelBigDecimal(scale = 2)
	private BigDecimal discountAmt;

	@ExcelField(name = "日期", sort = 4)
	@ExcelDate(format = "yyyy/MM/dd HH:mm")
	private Date createDate;

	@ExcelField(name = "会员", sort = 5, nullable = false)
	@ExcelBoolean(False = ExcelEnum.FALSE, True = ExcelEnum.TRUE)
	private Boolean isMember;

	public String getOrderNo() {
		return orderNo;
	}

	public void setOrderNo(String orderNo) {
		this.orderNo = orderNo;
	}

	public String getCustomerName() {
		return customerName;
	}

	public void setCustomerName(String customerName) {
		this.customerName = customerName;
	}

	public BigDecimal getOrderAmt() {
		return orderAmt;
	}

	public void setOrderAmt(BigDecimal orderAmt) {
		this.orderAmt = orderAmt;
	}

	public Date getCreateDate() {
		return createDate;
	}

	public void setCreateDate(Date createDate) {
		this.createDate = createDate;
	}

	public BigDecimal getDiscountAmt() {
		return discountAmt;
	}

	public void setDiscountAmt(BigDecimal discountAmt) {
		this.discountAmt = discountAmt;
	}

	public Boolean getIsMember() {
		return isMember;
	}

	public void setIsMember(Boolean isMember) {
		this.isMember = isMember;
	}

	@Override
	public String toString() {
		return "OrderForm [orderNo=" + orderNo + ", customerName=" + customerName + ", orderAmt=" + orderAmt
				+ ", discountAmt=" + discountAmt + ", createDate=" + createDate + ", isMember=" + isMember + "]";
	}
}
