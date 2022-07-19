package jiraApi;

import io.restassured.RestAssured;
import io.restassured.filter.session.SessionFilter;

import static io.restassured.RestAssured.given;

public class JiraTest {

	public static void main(String[] args)
	{
		RestAssured.baseURI="http//localhost:8080";
		
		//Login scenario
		
		SessionFilter session =new SessionFilter();
		String Response=given().header("Content-Type","application/json").body("").log().all().filter(session).when().post().asString();
	}
	
	
}
