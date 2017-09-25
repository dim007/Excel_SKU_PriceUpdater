//DIEGO MARTINEZ

import javax.swing.JFrame;


public class mainDriver {
	
	public static void main(String[] args) {
		guiMaker topFrame = new guiMaker();
		topFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		topFrame.pack();
		topFrame.setLocationRelativeTo(null);
		topFrame.setVisible(true);
		return;
	}

}
