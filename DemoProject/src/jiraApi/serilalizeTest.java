package jiraApi;

import io.restassured.RestAssured;
import io.restassured.response.Response;

import static io.restassured.RestAssured.given;

import java.util.ArrayList;
import java.util.List;

public class serilalizeTest {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Location l=new Location();
		AddPlace ap=new AddPlace();
		ap.setAccuracy(50);
		ap.setAddress("side layout");
		ap.setLanguage("HIndi");
		ap.setName("Robot");
		ap.setPhone_number("123456789");
		ap.setWebsite("http://google.com");
		List<String> myList=new ArrayList<String>();
		myList.add("park");
		myList.add("show");
		ap.setTypes(myList);
		l.setLatitude(-40.8560);
		l.setLongitude(90.9989);
		ap.setLocation(l);
		
		
		RestAssured.baseURI="https://rahulshettyacademy.com";
		Response res=given().log().all().
				queryParams("key","qaclick123").body(ap).when().post("/maps/api/place/add/json").then().assertThat().statusCode(200)
		.extract().response();
		
		String response=res.asString();
		System.out.println(response);

	}

}
