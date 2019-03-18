package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"html/template"
	"io/ioutil"
	"log"

	"github.com/tealeg/xlsx"
	"gopkg.in/gomail.v2"
)

type Student struct {
	FirstName   string `json:"FirstName"`
	GitHub      string `json:"GitHub"`
	LinkedIn    string `json:"LinkedIn"`
	UserName    string `json:"UserName"`
	Password    string `json:"Password"`
	EmailServer string `json:"EmailServer"`
	Port        int    `json:"Port"`
	Resume      string `json:"Resume"`
	Target      string `json:"Target"`
}
type SmtpTemplateData struct {
	EmployeeFirstName string
	EmployeeCompany   string
	StudentFirstName  string
	StudentGitHub     string
	StudentLinkedIn   string
}

type Employee struct {
	FirstName string
	Company   string
	EmailId   string
}

func (e *Employee) SendEmail(s Student) {
	var doc bytes.Buffer
	byteBuffer, err := ioutil.ReadFile("draft")
	if err != nil {
		fmt.Println("Failed to read draft")
	}
	emailTemplate := string(byteBuffer)
	t := template.New("emailTemplate")
	t, err = t.Parse(emailTemplate)
	if err != nil {
		log.Print("error trying to parse mail template")
	}
	context := &SmtpTemplateData{
		EmployeeFirstName: e.FirstName,
		StudentFirstName:  s.FirstName,
		EmployeeCompany:   e.Company,
		StudentGitHub:     s.GitHub,
		StudentLinkedIn:   s.LinkedIn,
	}
	err = t.Execute(&doc, context)
	if err != nil {
		log.Print("error trying to execute mail template")
	}
	msg := gomail.NewMessage()
	msg.SetHeader("From", s.UserName+"@gmail.com")
	msg.SetHeader("To", e.EmailId)
	msg.SetHeader("Subject", "Looking for "+s.Target+" Opportunities with "+e.Company)
	msg.SetBody("text/html", string(doc.Bytes()))
	msg.Attach(s.Resume)
	d := gomail.NewDialer("smtp.gmail.com", 587, s.UserName, s.Password)
	if err := d.DialAndSend(msg); err != nil {
		panic(err)
	}
}

func NewEmployee(row *xlsx.Row) *Employee {
	firstName := row.Cells[0].Value
	company := row.Cells[1].Value
	emailId := row.Cells[2].Value
	if firstName == "" || firstName == "FirstName" || company == "" || emailId == "" {
		return nil
	}
	return &Employee{FirstName: firstName, Company: company, EmailId: emailId}
}

func main() {
	studentInfo, err := ioutil.ReadFile("StudentInfo.json")
	if err != nil {
		fmt.Println(err)
		return
	}
	var student Student
	err = json.Unmarshal(studentInfo, &student)
	fmt.Println(student)
	xlFile, err := xlsx.OpenFile("contacts.xlsx")
	if err != nil {
		fmt.Println("Failed to read contacts file")
	}
	for _, sheet := range xlFile.Sheets {
		for _, row := range sheet.Rows {
			if len(row.Cells) > 0 {
				employee := NewEmployee(row)
				if employee != nil {
					employee.SendEmail(student)
				}
			}
		}
	}
}
