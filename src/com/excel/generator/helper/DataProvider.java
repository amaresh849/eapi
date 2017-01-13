package com.excel.generator.helper;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class DataProvider {

	public List<ProductDto> getProductData() {
		List<ProductDto> produDataList = new ArrayList<ProductDto>();

		ProductDto productDto1 = new ProductDto();
		productDto1.setProductId(1);
		productDto1.setProductName("A");
		productDto1.setProductPrice(111);
		productDto1.setQuantity(10);
		productDto1.setManufacturingDate(new Date());

		ProductDto productDto2 = new ProductDto();
		productDto2.setProductId(2);
		productDto2.setProductName("B");
		productDto2.setProductPrice(222);
		productDto2.setQuantity(10);
		productDto2.setManufacturingDate(new Date());

		ProductDto productDto3 = new ProductDto();
		productDto3.setProductId(3);
		productDto3.setProductName("C");
		productDto3.setProductPrice(333);
		productDto3.setQuantity(10);
		productDto3.setManufacturingDate(new Date());

		ProductDto productDto4 = new ProductDto();
		productDto4.setProductId(4);
		productDto4.setProductName("D");
		productDto4.setProductPrice(444);
		productDto4.setQuantity(10);
		productDto4.setManufacturingDate(new Date());

		produDataList.add(productDto1);
		produDataList.add(productDto2);
		produDataList.add(productDto3);
		produDataList.add(productDto4);
		return produDataList;
	}
}
