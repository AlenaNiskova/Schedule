package alena;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

public class Raspisanie extends JFrame {
    private JLabel timestart, timeend;
    private JButton okey;
    JTextField start, end;

    public Raspisanie() {
        super("Расписание");

        timestart = new JLabel("Первая уч. неделя:");
        timeend = new JLabel("Последняя уч. неделя:");
        start = new JTextField("");
        end = new JTextField("");
        okey = new JButton("Окей");

        JPanel panel = new JPanel();
        panel.setLayout(new GridLayout(3,2,5,10));
        panel.add(timestart);
        panel.add(timeend);
        panel.add(start);
        panel.add(end);
        panel.add(okey);

        add(panel);

        okey.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                
            }
        });

        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    }

    public static void main(String[] args) {
        Raspisanie rsp = new Raspisanie();
        rsp.setVisible(true);
        rsp.pack();
    }
}
