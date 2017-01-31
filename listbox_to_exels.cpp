#pragma once
#include <iostream>
#include <Windows.h>
namespace listbox {
	using namespace Excel;
	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace System::Diagnostics;
	using namespace System::IO;


	/// <summary>
	/// Summary for listboxTest
	/// </summary>
	public ref class listboxTest : public System::Windows::Forms::Form
	{
	public:
		listboxTest(void)
		{
			InitializeComponent();
			//
			//TODO: Add the constructor code here
			//
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~listboxTest()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^  button1;
	private: System::Windows::Forms::Button^  button2;
	private: System::Windows::Forms::Button^  button3;
	private: System::Windows::Forms::ListBox^  listBox1;
	private: System::Windows::Forms::DateTimePicker^  dateTimePicker1;
	protected:

	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		void InitializeComponent(void)
		{
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->button2 = (gcnew System::Windows::Forms::Button());
			this->button3 = (gcnew System::Windows::Forms::Button());
			this->listBox1 = (gcnew System::Windows::Forms::ListBox());
			this->dateTimePicker1 = (gcnew System::Windows::Forms::DateTimePicker());
			this->SuspendLayout();
			// 
			// button1
			// 
			this->button1->Location = System::Drawing::Point(39, 47);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(75, 23);
			this->button1->TabIndex = 0;
			this->button1->Text = L"Get Data";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &listboxTest::button1_Click);
			// 
			// button2
			// 
			this->button2->Location = System::Drawing::Point(39, 91);
			this->button2->Name = L"button2";
			this->button2->Size = System::Drawing::Size(75, 23);
			this->button2->TabIndex = 1;
			this->button2->Text = L"Clear Data";
			this->button2->UseVisualStyleBackColor = true;
			this->button2->Click += gcnew System::EventHandler(this, &listboxTest::button2_Click);
			// 
			// button3
			// 
			this->button3->Location = System::Drawing::Point(39, 132);
			this->button3->Name = L"button3";
			this->button3->Size = System::Drawing::Size(75, 23);
			this->button3->TabIndex = 2;
			this->button3->Text = L"Launch Exel";
			this->button3->UseVisualStyleBackColor = true;
			this->button3->Click += gcnew System::EventHandler(this, &listboxTest::button3_Click);
			// 
			// listBox1
			// 
			this->listBox1->FormattingEnabled = true;
			this->listBox1->ItemHeight = 12;
			this->listBox1->Location = System::Drawing::Point(170, 50);
			this->listBox1->Name = L"listBox1";
			this->listBox1->Size = System::Drawing::Size(450, 496);
			this->listBox1->TabIndex = 3;
			// 
			// dateTimePicker1
			// 
			dateTimePicker1->Format = DateTimePickerFormat::Custom;
			this->dateTimePicker1->CustomFormat = "yyyyMMdd";
			this->dateTimePicker1->Location = System::Drawing::Point(170, 12);
			this->dateTimePicker1->Name = L"dateTimePicker1";
			this->dateTimePicker1->Size = System::Drawing::Size(118, 19);
			this->dateTimePicker1->TabIndex = 4;
			// 
			// listboxTest
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 12);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(800, 600);
			this->Controls->Add(this->dateTimePicker1);
			this->Controls->Add(this->listBox1);
			this->Controls->Add(this->button3);
			this->Controls->Add(this->button2);
			this->Controls->Add(this->button1);
			this->Name = L"listboxTest";
			this->Text = L"listboxTest";
			this->ResumeLayout(false);

		}
#pragma endregion
	private: bool LoadCsvFile(String^ filePath) {
		try {
			StreamReader^ reader = gcnew StreamReader(filePath, System::Text::Encoding::GetEncoding("shift-jis"));
			//	reader->ReadLine();

			cli::array<wchar_t>^ separator = { ',' };
			String^ theDay = dateTimePicker1->Text;
			String^ selectDay = theDay->ToString();
			Debug::WriteLine(selectDay);
			String^ data;

			int count = 1;
			Excel::Application^ xls = gcnew Excel::ApplicationClass();
			Workbook^ wbook = xls->Workbooks->Add(Type::Missing);
			xls->Visible = true;
			Worksheet^ wSheet = static_cast<Worksheet^>(xls->ActiveSheet);
			wSheet->Name = "Active Sheet 1";

			while ((data = reader->ReadLine()) != nullptr) {
				cli::array<String^>^ split = data->Split(separator);
				int date = int::Parse(split[0]);
				int time = int::Parse(split[1]);
				int local = int::Parse(split[2]);
				String^ zeros = split[3];
				int person = int::Parse(split[4]);
				int area = int::Parse(split[5]);
				String^ nameInstruments = split[6];
				double valueOfInstruments = double::Parse(split[7]);

				if (selectDay == split[0]) {
					String^ note = String::Format("{0},{1},{2},{3},{4},{5},{6},{7}", split[0], split[1], split[2], split[3], split[4], split[5], split[6], split[7]);
					this->listBox1->Items->Add(note);
					putDataExels(wSheet,split[0],split[6],split[7], count);
					count++;
				}
			}
			
			reader->Close();

			String^ fname = "C:/Users/X220_M/Documents/Visual Studio 2015/Projects/listboxTest/listboxTest/data/test.xls";
			wbook->Application->DisplayAlerts = false;
			wbook->SaveAs(fname, XlFileFormat::xlWorkbookNormal, Type::Missing, Type::Missing, Type::Missing,
				Type::Missing, XlSaveAsAccessMode::xlNoChange, XlSaveConflictResolution::xlLocalSessionChanges,true,Type::Missing,Type::Missing,true);

			wbook->Close(false, Type::Missing, Type::Missing);
			xls->Quit();

			this->button1->Enabled = true;

			wbook->Application->DisplayAlerts = true;

		}
		catch (Exception^ ex) {
			return false;
		}

		
		return true;

	}



	private: System::Void button1_Click(System::Object^  sender, System::EventArgs^  e) {
		this->button1->Enabled = false;
		if (LoadCsvFile("./data/genbalog.txt") == true) {
		}
	}


	private: System::Void button2_Click(System::Object^  sender, System::EventArgs^  e) {
		this->button2->Enabled = false;
		this->listBox1->Items->Clear();
		this->button2->Enabled = true;
	}

		private: System::Void button3_Click(System::Object^  sender, System::EventArgs^  e) {
			Excel::Application^ xls = gcnew Excel::ApplicationClass();
			Workbook^ wbook = xls->Workbooks->Add(Type::Missing);
			xls->Visible = true;
			Worksheet^ wSheet = static_cast<Worksheet^>(xls->ActiveSheet);
			wSheet->Name = "Active Sheet 1";
			wSheet->Cells[1, 2] = "abc";// Cells[row,col]
		}

		 private: bool putDataExels(Worksheet^ wsName,String^ nameOfArray1,String^ nameOfArray2, String^ nameOfArray3,int rows) {

			 wsName->Cells[rows, 2] = nameOfArray1;// Cells[row,col]
			 wsName->Cells[rows, 3] = nameOfArray2;
			 wsName->Cells[rows, 4] = nameOfArray3;
			 
			 return true;
		 }

};
}
