/**intGIRLS 2019 Grading
 * Made by Lucinda Zhou
 * Using:
 * Apache (Excel files): https://poi.apache.org/
 * HMMT grading: https://www.hmmt.co/static/scoring-algorithm.pdf
 **/

package integirls;

public class Person {
	private int[] probs;
	private double score;
	private String name;
	
	public Person(String name1, int numProblems){
		name = name1;
		probs = new int[numProblems];
		score = 0;
	}
	
	public void setProb(int prob, int correct){
		probs[prob-1] = correct;
	}
	
	public int getProb(int prob){
		return probs[prob-1];
	}
	
	public void setScore(double newScore){
		score = newScore;
	}
	
	public double getScore(){
		return score;
	}
	
	public String getName(){
		return name;
	}
}
