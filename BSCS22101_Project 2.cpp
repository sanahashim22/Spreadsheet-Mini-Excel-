#include<windows.h>
#include<iostream>
#include<conio.h>
#include<iomanip>
#include <string>
#include<fstream>
#include <sstream>
#include<mmsystem.h>
using namespace std;

void getRowColbyLeftClick(int& rpos, int& cpos)
{
	HANDLE hInput = GetStdHandle(STD_INPUT_HANDLE);
	DWORD Events;
	INPUT_RECORD InputRecord;
	SetConsoleMode(hInput, ENABLE_PROCESSED_INPUT | ENABLE_MOUSE_INPUT | ENABLE_EXTENDED_FLAGS);
	do
	{
		ReadConsoleInput(hInput, &InputRecord, 1, &Events);
		if (InputRecord.Event.MouseEvent.dwButtonState == FROM_LEFT_1ST_BUTTON_PRESSED)
		{
			cpos = InputRecord.Event.MouseEvent.dwMousePosition.X;
			rpos = InputRecord.Event.MouseEvent.dwMousePosition.Y;
			break;
		}
	} while (true);
}
void gotoRowCol(int rpos, int cpos)
{
	COORD scrn;
	HANDLE hOuput = GetStdHandle(STD_OUTPUT_HANDLE);
	scrn.X = cpos;
	scrn.Y = rpos;
	SetConsoleCursorPosition(hOuput, scrn);
}
void color(int clr)
{
	SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), clr);
}

class myExcel {
	int cellWidth, cellHeight;
	class cell {
		string data;
		cell* right;
		cell* left;
		cell* up;
		cell* down;
		friend class myExcel;
	public:
		cell(string da = "", cell* l = nullptr, cell* r = nullptr, cell* u = nullptr, cell* d = nullptr) {
			data = da;
			left = l;
			right = r;
			up = u;
			down = d;
		}
	};
	int r_size = 0, c_size = 0;
	int r_curr = 0, c_curr = 0;
	cell* curr;
	cell* head;
	cell* tail;
	cell* rangestart = nullptr;
	cell* rangeend = nullptr;
	int cs, rs, re, ce;
public:
	friend class cell;
	myExcel() {
		cellWidth = 10;
		cellHeight = 3;
		r_size = 1;
		c_size = 1;
		head = new cell();
		tail = head;
		curr = head;
		r_curr = 0;
		c_curr = 0;
		for (int i = 0; i < 4; i++) {
			insertColRight();
			curr = curr->right;
		}
		curr = head;
		for (int i = 0; i < 4; i++) {
			insertRowBelow();
			curr = curr->down;
		}
		curr = head;
		printGrid(); 
		printData();
		printcellborder(r_curr, c_curr, 4);
		printButton();
		inputkeyboard();
	}
	void printData() {
		color(7);
		cell* temp = head;
		cell* storCurr = curr;
		int storCurrRow = r_curr;
		int storCurrCol = c_curr;
		r_curr = 0;
		while (temp != nullptr) {
			c_curr = 0;
			curr = temp;
			while (curr != nullptr) {
				printincell();
				curr = curr->right;
				c_curr++;

			}
			cout << endl;
			temp = temp->down;
			r_curr++;
		}
		r_curr = storCurrRow;
		c_curr = storCurrCol;
		curr = storCurr;
	}
	void printincell() {
		color(7);
		int r = r_curr * cellHeight + cellHeight / 2;
		int c = c_curr * cellWidth + cellWidth / 2;
		gotoRowCol(r, c);
		cout << curr->data;
	}
	void printincell(int row, int col) {
		color(7);
		int r = row * cellHeight + cellHeight / 2;
		int c = col * cellWidth + cellWidth / 2;
		gotoRowCol(r, c);
		cout << curr->data;
	}
	void eraseincell(int row, int col) {
		color(7);
		int r = row * cellHeight + cellHeight / 2;
		int c = col * cellWidth + cellWidth / 2;
		gotoRowCol(r, c);
		curr->data = "";
		cout << curr->data;
	}
	void printcellborder(int row, int col, int clr = 7) {
		color(clr);
		int r_curr1 = row;
		int c_curr1 = col;
		cellWidth = 10;
		cellHeight = 3;
		int w = 0, h = 0;
		int r = r_curr1 * cellHeight;
		int c = c_curr1 * cellWidth;
		gotoRowCol(r, c);
		for (w = 0; w < cellWidth; w++) {
			cout << "*";
		}
		w = 0;
		r = (r_curr1 * cellHeight) + cellHeight;
		gotoRowCol(r, c);
		for (w = 0; w <= cellWidth; w++) {
			cout << "*";
		}
		r = r_curr1 * cellHeight;
		c = c_curr1 * cellWidth;
		for (h = 0; h < cellHeight; h++) {
			gotoRowCol(r + h, c);
			cout << "*";
		}
		c = (c_curr1 * cellWidth) + cellWidth;
		r = r_curr1 * cellHeight;
		for (h = 0; h < cellHeight; h++) {
			gotoRowCol(r + h, c);
			cout << "*";
		}
	}
	void printGrid() {
		for (int i = 0; i < r_size; i++) {
			for (int c = 0; c < c_size; c++) {
				printcellborder(i, c);
			}
		}
	}
	void insertColRight() {
		cell* temp = curr;
		while (curr->up != nullptr) {
			curr = curr->up;
		}
		while (curr != nullptr) {
			insertCelRight(curr);
			curr = curr->down;
		}
		curr = temp;
		c_size++;
	}
	cell* insertCelRight(cell*& c) {
		cell* temp = new cell();
		temp->left = c;
		if (c->right != nullptr) {
			temp->right = c->right;
			temp->right->left = temp;
		}
		c->right = temp;
		if (c->up != nullptr && c->up->right != nullptr) {
			temp->up = c->up->right;
			c->up->right->down = temp;
		}
		if (c->down != nullptr && c->down->right != nullptr) {
			temp->down = c->down->right;
			c->down->right->up = temp;
		}
		return temp;
	}
	void insertRowBelow() {
		cell* temp = curr;
		while (curr->left != nullptr) {
			curr = curr->left;
		}
		while (curr != nullptr) {
			insertCellBelow(curr);
			curr = curr->right;
		}
		curr = temp;
		r_size++;
	}
	cell* insertCellBelow(cell*& c) {
		cell* temp = new cell();
		temp->up = c;
		if (c->down != nullptr) {
			temp->down = c->down;
			temp->down->up = temp;
		}
		curr->down = temp;
		if (c->left != nullptr && c->left->down != nullptr) {
			temp->left = c->left->down;
			c->left->down->right = temp;
		}
		if (c->right != nullptr && c->right->down != nullptr) {
			temp->right = c->right->down;
			c->right->down->left = temp;
		}
		return temp;
	}
	cell* insertCellAbove(cell*& c) {
		cell* temp = new cell();
		temp->down = c;
		if (c->up != nullptr) {
			temp->up = c->up;
			temp->up->down = temp;
		}
		c->up = temp;
		if (c->left != nullptr && c->left->up != nullptr) {
			temp->left = c->left->up;
			c->left->up->right = temp;
		}
		if (c->right != nullptr && c->right->up != nullptr) {
			temp->right = c->right->up;
			c->right->up->left = temp;
		}
		return temp;
	}
	void insertRowAbove() {
		cell* temp = curr;
		while (curr->left != nullptr) {
			curr = curr->left;
		}
		while (curr != nullptr) {
			insertCellAbove(curr);
			curr = curr->right;
		}
		curr = temp;
		r_size++;
	}
	cell* insertCellLeft(cell*& c) {
		cell* temp = new cell();
		temp->right = c;
		if (c->left != nullptr) {
			temp->left = c->left;
			temp->left->right = temp;
		}
		c->left = temp;
		if (c->up != nullptr && c->up->left != nullptr) {
			temp->up = c->up->left;
			c->up->left->down = temp;
		}
		if (c->down != nullptr && c->down->left != nullptr) {
			temp->down = c->down->left;
			c->down->left->up = temp;
		}
		return temp;
	}
	void insertColLeft() {
		cell* temp = curr;
		while (curr->up != nullptr) {
			curr = curr->up;
		}
		while (curr != nullptr) {
			insertCellLeft(curr);
			curr = curr->down;
		}
		curr = temp;
		c_size++;
	}

	void clearRow() {
		cell* temp = curr;
		temp->data = curr->data;
		int r = c_curr;
		while (curr->left != nullptr) {
			curr = curr->left;
			c_curr--;
		}
		while (curr->right != nullptr) {
			curr->data = "";
			//printincell(r_curr,c_curr);
			curr = curr->right;
			c_curr++;
		}
		curr = temp;
		curr->data = "";
		c_curr = r;
		system("cls");
		printGrid();
		printData();
		printcellborder(r_curr, c_curr, 4);
		printButton();
	}
	void clearCol() {
		cell* temp = curr;
		temp->data = curr->data;
		int r = r_curr;
		while (curr->up != nullptr) {
			curr = curr->up;
			r_curr--;
		}
		while (curr->down != nullptr) {
			curr->data = "";
			curr = curr->down;
			r_curr++;
		}
		curr = temp;
		curr->data = "";
		r_curr = r;
		system("cls");
		printGrid();
		printData();
		printcellborder(r_curr, c_curr, 4);
		printButton();
	}
	void saveToFile(const string& filename) {
		ofstream wrt(filename);
		if (wrt.is_open()) {
			cell* temp = head;
			for (int r = 0; r < r_size; r++) {
				cell* temp2 = temp;
				for (int c = 0; c < c_size; c++) {
					wrt << temp->data;
					if (c < c_size - 1) {
						wrt << ",";
					}
					temp = temp->right;
				}
				wrt << "\n";
				temp = temp2->down;
			}
			wrt.close();
		}
	}


	void loadFromFile(const string& filename) {
		ifstream file(filename);
		if (file.is_open()) {
			clearData();
			string line;
			int r = 0;
			while (getline(file, line)) {
				stringstream ss(line);
				string cellData;
				int c = 0;
				while (getline(ss, cellData, ',')) {
					if (r < r_size && c < c_size) {
						cell* temp = head;
						for (int i = 0; i < r; i++)
							temp = temp->down;
						for (int j = 0; j < c; j++)
							temp = temp->right;
						temp->data = cellData;
					}
					c++;
				}
				r++;
			}
			file.close();
			system("cls");
			printGrid();
			printData();
			printButton();
			printcellborder(r_curr, c_curr, 4);
		}
	}



	void clearData() {
		color(7);
		cell* temp = head;
		cell* storCurr = curr;
		int storCurrRow = r_curr;
		int storCurrCol = c_curr;
		r_curr = 0;
		while (temp != nullptr) {
			c_curr = 0;
			curr = temp;
			while (curr != nullptr) {
				curr->data = "";
				//curr->data = "";
				//printincell();
				curr = curr->right;
				c_curr++;

			}
			temp = temp->down;
			r_curr++;
		}
		r_curr = storCurrRow;
		c_curr = storCurrCol;
		curr = storCurr;
		system("cls");
		printGrid();
		printData();
		printcellborder(r_curr, c_curr, 4);
		printButton();
	}
	void delCol() {
		if (c_size <= 1) return;
		if (curr->right != nullptr)
			curr->right->left = curr->left;
		if (curr->left != nullptr)
			curr->left->right = curr->right;
		if (curr->right != nullptr)
			curr = curr->right;
		else
			curr = head;
		c_size--;
		system("cls");
		printGrid();
		printData();
		printcellborder(r_curr, c_curr, 4);
		printButton();
	}
	void delRow() {
		if (r_size <= 1) return;
		while (curr->left != nullptr)
			curr = curr->left;
		while (curr->right != nullptr) {
			if (curr->up != nullptr) {
				curr->up->down = curr->down;
				curr->up->down->data = curr->down->data;
			}
			if (curr->down != nullptr) {
				curr->down->up = curr->up;
				curr->down->up->data = curr->up->data;
			}
			if (curr->down != nullptr) {
				curr = curr->down;
				curr->data = curr->down->data;
			}
			else
				curr = head;
			curr = curr->right;
		}

		r_size--;
		system("cls");
		printGrid();
		printData();
		printcellborder(r_curr, c_curr, 4);
		printButton();
	}
	void inputkeyboard() {
		string d;
		while (1) {
			char k = _getch(); 
			int ascii = k;
			if (ascii == 27) {  //for ESC
				break;
			}
			else if (ascii == -32) {
				char c = _getch();
				ascii = c;
				if (ascii == 72) {  //up == w
					if (curr->up == nullptr) {
						insertRowAbove();
						printGrid();
					}
					printcellborder(r_curr, c_curr, 7);
					curr = curr->up;
					r_curr--;
					printcellborder(r_curr, c_curr, 4);
				}
				else if (ascii == 75) {  //left == a
					if (curr->left == nullptr) {
						insertColLeft();
						printGrid();
					}
					printcellborder(r_curr, c_curr, 7);
					curr = curr->left;
					c_curr--;
					printcellborder(r_curr, c_curr, 4);
				}
				else if (ascii == 77) {  //right  == d
					if (curr->right == nullptr) {
						insertColRight();
						printGrid();
					}
					printcellborder(r_curr, c_curr, 7);
					curr = curr->right;
					c_curr++;
					printcellborder(r_curr, c_curr, 4);
				}
				else if (ascii == 80) {  //down == s
					if (curr->down == nullptr) {
						insertRowBelow();
						printGrid();
					}
					printcellborder(r_curr, c_curr, 7);
					curr = curr->down;
					r_curr++;
					printcellborder(r_curr, c_curr, 4);
				}
			}
			
		
			//else if(ascii==100){  //right  == d
			//	if (curr->right == nullptr) {
			//		insertColRight();
			//		printGrid();
			//	}
			//	printcellborder(r_curr, c_curr, 7); 
			//	curr = curr->right;
			//	c_curr++;

			//	printcellborder(r_curr, c_curr, 4);

			//}
			//else if (ascii == 97) {  //left == a
			//	if (curr->left == nullptr) {
			//		insertColLeft();
			//		printGrid();
			//	}
			//	printcellborder(r_curr, c_curr, 7);
			//	curr = curr->left;
			//	c_curr--;
			//	printcellborder(r_curr, c_curr, 4);

			//}
			//else if (ascii == 119) {  //up == w
			//	if (curr->up == nullptr) {
			//		insertRowAbove();
			//		printGrid();
			//	}
			//	printcellborder(r_curr, c_curr, 7);
			//	curr = curr->up;
			//	r_curr--;
			//	printcellborder(r_curr, c_curr, 4);

			//}
			//else if (ascii == 115) {  //down == s
			//	if (curr->down == nullptr) {
			//		insertRowBelow();
			//		printGrid();
			//	}
			//	printcellborder(r_curr, c_curr, 7);
			//	curr = curr->down;
			//	r_curr++;
			//	printcellborder(r_curr, c_curr, 4);

			//}
			else if (ascii == 105) { //insert row col by pressing i
				k = _getch();
				ascii = k;
				if (ascii == 115) { //s for row
					insertRowBelow();
				}
				if (ascii == 100) { //d for col
					insertColRight();
				}
				if (ascii == 97) { //a for col left
					insertColLeft();
				}
				if (ascii == 119) { //w for row above
					insertRowAbove();
				}
				printGrid();
				printData();
				printcellborder(r_curr, c_curr, 4);
			}
			else if (ascii == 8) { //backspace for clearing
				k = _getch();
				ascii = k;
				if (ascii == 114) { //r for current row
					clearRow();
				}
				if (ascii == 99) { //c for col current
					clearCol();
				}
				printGrid();
				printData();
				printcellborder(r_curr, c_curr, 4);
			}
			/*else if (ascii == 120) {
				delRow();
			}
			else if (ascii == 121) {
				delCol();
			}*/
			else if (ascii == 3) { //cntrl+c for copying data
				d = curr->data;
			}
			else if (ascii == 22) {//cntrl+v to paste data
				curr->data = d;
				printincell();
			}
			else if (ascii == 24) { //cntrl+x for cut data
				d = curr->data;
				curr->data = "";
				system("cls");
				printGrid();
				printData();
				printcellborder(r_curr, c_curr, 4);
				printButton();
			}
			else if (ascii == 102) { //f to perform function
				functions();
			}
			//else if (ascii == 110) { // insert row col by rght shift pressing n
			//	k = _getch();
			//	ascii = k;
			//	if (ascii == 114) { // r for row
			//		InsertCellByDownShift();
			//	}
			//	if (ascii == 99) { // c for col
			//		InsertCellByRightShift();
			//	}
			//	printGrid();
			//	printData();
			//	printcellborder(r_curr, c_curr, 4);
			//}
			//else if (ascii == 121) { // y
			//	DeleteCellByLeftShift();
			//	printGrid();
			//	printData();
			//	printcellborder(r_curr, c_curr, 4);
			//}
			//else if (ascii == 120) { // x
			//	DeleteCellByUpShift();
			//	printGrid();
			//	printData();
			//	printcellborder(r_curr, c_curr, 4);
			//}
			/*else {
				char ch = ascii;
				curr->data = ch; 
				printincell();
			}*/
			else
			{
				if (isprint(k) || k == ' ') {
					if (curr->data.size() < 4) {
						curr->data += k;
					}
					printincell();
				}
}

		}
	}
	//for char data type 
	/*int sum() {

		int sum = 0;
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				sum += temp->data - '0';
				temp = temp->right;
			}
			temp = temp2->down;
		}
		return sum;
	}
	int average() {
		int sum = 0;
		int count = 0;
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				sum += temp->data - '0';
				count++;
				temp = temp->right;
			}
			temp = temp2->down;
		}
		return sum/count;
	}
	int count() {
		int count = 0;
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				if (temp->data != ' ' && temp->data !='\0') {
					count++;
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		return count;
	}*/
	//for char data type
	/*int Min() {
		int min = INT_MAX;
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				if (temp->data != " " && temp->data != '\0') {
					int val = temp->data - '0';
					if (val < min) {
						min = val;
					}
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		return min;
	}
	int Max() {
		int max = INT_MIN; 
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				if (temp->data != ' ' && temp->data != '\0') {
					int val = temp->data - '0';
					if (val > max) {
						max = val;
					}
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		return max;
	}
	*/
	int sum() {

		int sum = 0;
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				if (temp->data != " " && temp->data != "") {
					sum += stoi(temp->data);
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		return sum;
	}
	int average() {
		int sum = 0;
		int count = 0;
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				if (temp->data != " " && temp->data != "") {
					sum += stoi(temp->data);
					count++;
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		if (count == 0) {
			return 0; 
		}
		return sum / count;
	}
	int count() {
		int count = 0;
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				if (temp->data != " " && temp->data != "") {
					count++;
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		return count;
	}
	int Min() {
		int min = INT_MAX; 
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				if (temp->data != " " && !temp->data.empty()) {
					int val = stoi(temp->data);
					if (val < min) {
						min = val;
					}
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		return min;
	}

	int Max() {
		int max = INT_MIN; 
		cell* temp = rangestart;
		int rowiterations = re - rs;
		int coliterations = ce - cs;
		for (int r = 0; r <= rowiterations; r++) {
			cell* temp2 = temp;
			for (int c = 0; c <= coliterations; c++) {
				if (temp->data != " " && !temp->data.empty()) {
					int val = stoi(temp->data);
					if (val > max) {
						max = val;
					}
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		return max;
	}
	void functions() {
		int rpos, cpos;
		getRowColbyLeftClick(rpos, cpos);
		rpos /= cellHeight;
		cpos /= cellWidth;
		if (rpos == 1 && cpos == 10) {
			printcellborder(1, 10, 4);
			selectrange();
			inputkeyboard(); 
			curr->data = to_string(sum());
			printincell();
			printcellborder(1, 10, 7);
		} 
		if (rpos == 3 && cpos == 10) {
			printcellborder(3, 10, 4);
			selectrange();
			inputkeyboard();
			curr->data = to_string(average());
			printincell();
			printcellborder(3, 10, 7);
		}
		if (rpos == 5 && cpos == 10) {
			printcellborder(5, 10, 4);
			selectrange();
			inputkeyboard();
			curr->data = to_string(count());
			printincell();
			printcellborder(5, 10, 7);
		}
		if (rpos == 7 && cpos == 10) {
			printcellborder(7, 10, 4);
			selectrange();
			inputkeyboard();
			curr->data = to_string(Min());
			printincell();
			printcellborder(7, 10, 7);
		}
		if (rpos == 9 && cpos == 10) {
			printcellborder(9, 10, 4);
			selectrange();
			inputkeyboard();
			curr->data = to_string(Max());
			printincell();
			printcellborder(9, 10, 7);
		}

		if (rpos == 1 && cpos == 12) {
			printcellborder(1, 12, 4);
			clearData();
			printcellborder(1, 12, 7);
		}

		if (rpos == 3 && cpos == 12) {
			printcellborder(3, 12, 4);
			saveToFile("sl.txt");
			printincell();
			printcellborder(3, 12, 7);
		}
		if (rpos == 5 && cpos == 12) {
			printcellborder(5, 12, 4);
			loadFromFile("sl.txt");
			printcellborder(5, 12, 7);
		}
	}
	void selectrange() {
		int ascii = 0; int selected = 0;
		while (selected < 2) {
			char k = _getch();
			ascii = k;
			if (ascii == -32) { 
				char c = _getch();
				ascii = c;
				if (ascii == 77) {  //right d
					if (curr->right != nullptr) {
						printcellborder(r_curr, c_curr, 7);
						curr = curr->right;
						c_curr++;
						printcellborder(r_curr, c_curr, 4);
					}
				}
				else if (ascii == 75) {  //left == a
					printcellborder(r_curr, c_curr, 7);
					if (curr->left != nullptr) {
						curr = curr->left;
						c_curr--;
					}
					printcellborder(r_curr, c_curr, 4);
				}
				else if (ascii == 72) {  //up == w
					printcellborder(r_curr, c_curr, 7);
					if (curr->up != nullptr) {
						curr = curr->up;
						r_curr--;
					}
					printcellborder(r_curr, c_curr, 4);

				}
				else if (ascii == 80) {  //down == s
					printcellborder(r_curr, c_curr, 7);
					if (curr->down != nullptr) {
						curr = curr->down;
						r_curr++;
					}
					printcellborder(r_curr, c_curr, 4);

				}
			}
			else if (ascii == 32) {  //space bar for selection of cell in range
				if (selected == 0) {
					rangestart = curr;
					rs = r_curr;
					cs = c_curr; 
					printincell(12,10);
				}
				if (selected == 1) {
					rangeend = curr;
					re = r_curr;
					ce = c_curr;
					printincell(14, 10);
				}
				selected++;
			}
			else {
				char ch = ascii;
				curr->data = ch;
				printincell();
			}
		}
	}
	void printButton() {
		int r = 1, c = 10; 
		printcellborder(1, 10); 
		r = r * cellHeight + 2;
		c = c * cellWidth + 5;
		gotoRowCol(r, c); 
		cout << "Sum";

		printcellborder(3, 10);
		r = r +2* cellHeight;
		gotoRowCol(r, c);
		cout << "Avg";

		printcellborder(5, 10);
		r = r +2* cellHeight;
		gotoRowCol(r, c-1);
		cout << "Count";

		printcellborder(7, 10);
		r = r + 2*cellHeight;
		gotoRowCol(r, c);
		cout << "Min";

		printcellborder(9, 10);
		r = r + 2*cellHeight;
		gotoRowCol(r, c );
		cout << "Max";







		r = 1, c = 12;
		printcellborder(1, 12);
		r = r * cellHeight + 2;
		c = c * cellWidth + 5;
		gotoRowCol(r, c);
		cout << "Clear";

		printcellborder(3, 12);
		r = r + 2 * cellHeight;
		gotoRowCol(r, c);
		cout << "Save";

		printcellborder(5, 12);
		r = r + 2 * cellHeight;
		gotoRowCol(r, c - 1);
		cout << "Load";


		printcellborder(12, 10);
		printcellborder(14, 10);

	}
	

	//void DeleteColumn() {
	//	if (curr->left != nullptr) {
	//		cell* temp = curr;
	//		while (temp->up != nullptr) {
	//			temp = temp->up;
	//		}
	//		while (temp != nullptr) {
	//			cell* toDelete = temp;
	//			temp = temp->down;
	//			if (toDelete->left != nullptr) {
	//				toDelete->left->right = toDelete->right;
	//			}
	//			if (toDelete->right != nullptr) {
	//				toDelete->right->left = toDelete->left;
	//			}
	//			delete toDelete;
	//		}
	//		c_curr--;
	//		c_size--;
	//	}
	//}

	//void DeleteRow() {
	//	if (curr->up != nullptr) {
	//		cell* temp = curr;
	//		while (temp->left != nullptr) {
	//			temp = temp->left;
	//		}
	//		while (temp != nullptr) {
	//			cell* toDelete = temp;
	//			temp = temp->right;
	//			if (toDelete->up != nullptr) {
	//				toDelete->up->down = toDelete->down;
	//			}
	//			if (toDelete->down != nullptr) {
	//				toDelete->down->up = toDelete->up;
	//			}
	//			delete toDelete;
	//		}
	//		r_curr--;
	//		r_size--;
	//	}
	//}
	//
	/*
	void DeleteCellByLeftShift() {
		if (curr->left != nullptr) {
			cell* temp = curr->left;
			curr->left = temp->left;
			if (temp->left != nullptr) {
				temp->left->right = curr;
			}
			delete temp;
			curr = curr->left;
			c_curr--;
			c_size--;
		}
	}

	void DeleteCellByUpShift() {
		if (curr->up != nullptr) {
			cell* temp = curr->up;
			curr->up = temp->up;
			if (temp->up != nullptr) {
				temp->up->down = curr;
			}
			delete temp;
			curr = curr->up;
			r_curr--;
			r_size--;
		}
	}


	void InsertCellByRightShift() {
		if (curr->right == nullptr) {
			insertColRight();
		}
		cell* temp = curr->right;
		curr->right = temp->right;
		if (temp->right != nullptr) {
			temp->right->left = curr;
		}
		temp->right = curr;
		curr->left = temp;
		curr = curr->right;
		c_curr++;
		c_size++;
	}

	void InsertCellByDownShift() {
		if (curr->down == nullptr) {
			insertRowBelow();
		}
		cell* temp = curr->down;
		curr->down = temp->down;
		if (temp->down != nullptr) {
			temp->down->up = curr;
		}
		temp->down = curr;
		curr->up = temp;
		curr = curr->down;
		r_curr++;
		r_size++;
	}*/

	
};






int main() {
	//sndPlaySound(TEXT("new.wav"), SND_ASYNC);
	myExcel e;
	//cout << endl<<endl<<endl;
	gotoRowCol(50, 0);
	//system("pause");
	return 0;
}
