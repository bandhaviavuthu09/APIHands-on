package jiraApi;

import io.restassured.RestAssured;
import io.restassured.builder.RequestSpecBuilder;
import io.restassured.builder.ResponseSpecBuilder;
import io.restassured.http.ContentType;
import io.restassured.response.Response;
import io.restassured.specification.RequestSpecification;
import io.restassured.specification.ResponseSpecification;

import static io.restassured.RestAssured.given;

import java.util.ArrayList;
import java.util.List;

public class SpecBuilderTest {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		RestAssured.baseURI="https://rahulshettyacademy.com";
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
				
				 RequestSpecification req =new RequestSpecBuilder().setBaseUri("https://rahulshettyacademy.com").addQueryParam("key", "qaclick123")
						 .setContentType(ContentType.JSON).build();
				
				ResponseSpecification resspec=new ResponseSpecBuilder().expectStatusCode(200).expectContentType(ContentType.JSON).build();
				RequestSpecification res=given().spec(req).body(ap);
				
				Response response=res.when().post("/maps/api/place/add/json").then().spec(resspec).extract().response();
				
//				Response res=given().log().all().
//						queryParams("key","qaclick123").body(ap).when().post("/maps/api/place/add/json").then().assertThat().statusCode(200)
//				.extract().response();
				
				String responseSt=response.asString();
				System.out.println(response);


	}

}
