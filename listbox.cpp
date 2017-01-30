#pragma once
#include <string>
#include <iostream>
#include <Windows.h>

using namespace std;

namespace BarCodeDate {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace System::Diagnostics;
	using namespace System::IO;

	/// <summary>
	/// Summary for MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
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
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}

	private: System::Windows::Forms::Button^  button1;
	private: System::Windows::Forms::ListBox^  listBox1;
	private: System::Windows::Forms::DateTimePicker^  dateTimePicker1;
	private: System::Windows::Forms::Button^  button2;
	private: System::Windows::Forms::Button^  button3;
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
			this->listBox1 = (gcnew System::Windows::Forms::ListBox());
			this->dateTimePicker1 = (gcnew System::Windows::Forms::DateTimePicker());
			this->button2 = (gcnew System::Windows::Forms::Button());
			this->button3 = (gcnew System::Windows::Forms::Button());
			this->SuspendLayout();
			// 
			// button1
			// 
			this->button1->Font = (gcnew System::Drawing::Font("MS UI Gothic", 11.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(128)));
			this->button1->Location = System::Drawing::Point(18, 131);
			this->button1->Margin = System::Windows::Forms::Padding(4);
			this->button1->Name = "button1";
			this->button1->Size = System::Drawing::Size(140, 40);
			this->button1->TabIndex = 1;
			this->button1->Text = "Get Data";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click);
			// 
			// listBox1
			// 
			this->listBox1->Font = (gcnew System::Drawing::Font("MS UI Gothic", 11.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(128)));
			this->listBox1->FormattingEnabled = true;
			this->listBox1->ItemHeight = 15;
			this->listBox1->Location = System::Drawing::Point(345, 42);
			this->listBox1->Margin = System::Windows::Forms::Padding(4);
			this->listBox1->Name = "listBox1";
			this->listBox1->Size = System::Drawing::Size(500, 589);
			this->listBox1->TabIndex = 2;


			// 
			// dateTimePicker1
			// 
			dateTimePicker1->Format = DateTimePickerFormat::Custom;
			this->dateTimePicker1->CustomFormat = "yyyyMMdd";
			this->dateTimePicker1->Font = (gcnew System::Drawing::Font("MS UI Gothic", 11.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(128)));
			this->dateTimePicker1->Location = System::Drawing::Point(18, 42);
			this->dateTimePicker1->Margin = System::Windows::Forms::Padding(4);
			this->dateTimePicker1->Name = "dateTimePicker1";
			this->dateTimePicker1->Size = System::Drawing::Size(298, 22);
			this->dateTimePicker1->TabIndex = 3;
			// 
			// button2
			// 
			this->button2->Font = (gcnew System::Drawing::Font("MS UI Gothic", 11.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(128)));
			this->button2->Location = System::Drawing::Point(18, 210);
			this->button2->Margin = System::Windows::Forms::Padding(4);
			this->button2->Name = L"button2";
			this->button2->Size = System::Drawing::Size(140, 40);
			this->button2->TabIndex = 4;
			this->button2->Text = L"Clear Data";
			this->button2->UseVisualStyleBackColor = true;
			this->button2->Click += gcnew System::EventHandler(this, &MyForm::button2_Click);
			// 
			// button3
			// 
			this->button3->Font = (gcnew System::Drawing::Font("MS UI Gothic", 11.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(128)));
			this->button3->Location = System::Drawing::Point(18, 290);
			this->button3->Name = "button3";
			this->button3->Size = System::Drawing::Size(140, 40);
			this->button3->TabIndex = 5;
			this->button3->Text = "Open Exel";
			this->button3->UseVisualStyleBackColor = true;
			this->button3->Click += gcnew System::EventHandler(this, &MyForm::button3_Click);
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(9, 15);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(980, 720);
			this->Controls->Add(this->button3);
			this->Controls->Add(this->button2);
			this->Controls->Add(this->dateTimePicker1);
			this->Controls->Add(this->listBox1);
			this->Controls->Add(this->button1);
			this->Font = (gcnew System::Drawing::Font("MS UI Gothic", 11.25F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(128)));
			this->Margin = System::Windows::Forms::Padding(4);
			this->Name = "MyForm";
			this->Text = "MyForm";
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
				}
			}

			reader->Close();
			this->button1->Enabled = true;
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
	}
	
};
}
