// src/App.jsx
import React from "react";
import { Formik, Form, Field, ErrorMessage } from "formik";
import * as Yup from "yup";
import SignatureCanvas from "react-signature-canvas";
import textTermsAndConditions from "./TextContent";
import axios from "axios";

const yourClientId = import.meta.env.VITE_AZURE_CLIENT_ID;
const yourClientSecret = import.meta.env.VITE_AZURE_CLIENT_SECRET;
const yourRedirectUri = import.meta.env.VITE_AZURE_REDIRECT_URL;
const yourSharePointSiteUrl = import.meta.env.VITE_SHAREPOINT_SITE_URL;
const yourListName = import.meta.env.VITE_SHAREPOINT_LIST_NAME;
const yourAccessToken = import.meta.env.VITE_ACCESS_TOKEN;

const validationSchema = Yup.object().shape({
  address: Yup.string().required("Address is required"),
  city: Yup.string().required("City is required"),
  dateOfBirth: Yup.date().required("Date of Birth is required"), 
  firstName: Yup.string().required("First Name is required"),
  lastName: Yup.string().required("Last Name is required"),
  phoneNumber: Yup.string().required("Phone Number is required"),
  postalCode: Yup.string().required("Postal Code is required"),
  province: Yup.string().required("Province is required"),
  publicHealthCardNumber: Yup.string().required("Public Health Card Number is required"),
  // signature: Yup.string().required("Signature is required"),
});

const initialValues = {
  address: "",
  city: "",
  dateOfBirth: "",
  firstName: "",
  lastName: "",
  email: "",
  phoneNumber: "",
  postalCode:"",
  province: "",
  publicHealthCardNumber: "",
  signature: "",
};



const onSubmit = async (values, { setSubmitting }) => {
  try {

    
    // Make HTTP POST request to SharePoint API
    const response = await axios.post(
      `${yourSharePointSiteUrl}/_api/web/lists/getbytitle('${yourListName}')/items`,
      {
        // Map form values to SharePoint list columns
        Title: values.firstName,
        Address: values.address,
        SignatureData: values.signatureData, // Adjust this based on your SharePoint column name
        // Map other fields accordingly
      },
      {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'Authorization': `Bearer ${yourAccessToken}`,
          // Include Azure AD app details
          'client_id': yourClientId,
          'client_secret': yourClientSecret,
          'redirect_uri': yourRedirectUri,
        },
      }
    );

    console.log('SharePoint API Response:', response.data);
  } catch (error) {
    console.error('Error submitting to SharePoint:', error);
  } finally {
    setSubmitting(false);
  }
};

function App() {

  const signaturePad = React.useRef();

  return (
    <div className="min-h-screen flex items-center justify-center bg-gray-100">
    <Formik
      initialValues={initialValues}
      validationSchema={validationSchema}
      onSubmit={onSubmit}
    >
      <Form className="w-full max-w-3xl bg-white p-8 rounded-md shadow-md">
        <h1 className="text-2xl font-bold mb-10 text-center">VCH Patient Consent Form</h1>

        {/* First Name and Last Name in a flex container */}
        <div className="flex mb-4">
          <div className="w-1/2 pr-2">
            <label htmlFor="firstName" className="block text-sm font-bold mb-2">
              First Name:
            </label>
            <Field
              type="text"
              id="firstName"
              name="firstName"
              className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
            />
            <ErrorMessage name="firstName" component="p" className="text-red-500 text-xs mt-1" />
          </div>

          <div className="w-1/2 pl-2">
            <label htmlFor="lastName" className="block text-sm font-bold mb-2">
              Last Name:
            </label>
            <Field
              type="text"
              id="lastName"
              name="lastName"
              className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
            />
            <ErrorMessage name="lastName" component="p" className="text-red-500 text-xs mt-1" />
          </div>
        </div>


         {/* Public Health Card Number */}
         <div className="mb-4">
            <label htmlFor="publicHealthCardNumber" className="block text-sm font-bold mb-2">
              Public Health Card Number:
            </label>
            <Field
              type="text"
              id="publicHealthCardNumber"
              name="publicHealthCardNumber"
              className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
            />
            <ErrorMessage
              name="publicHealthCardNumber"
              component="p"
              className="text-red-500 text-xs mt-1"
            />
          </div>

          {/* Date of Birth */}
          <div className="mb-4">
            <label htmlFor="dateOfBirth" className="block text-sm font-bold mb-2">
              Date of Birth:
            </label>
            <Field
              type="date"
              id="dateOfBirth"
              name="dateOfBirth"
              className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
            />
            <ErrorMessage name="dateOfBirth" component="p" className="text-red-500 text-xs mt-1" />
          </div>

         {/* Phone Number */}
         <div className="mb-4">
            <label htmlFor="phoneNumber" className="block text-sm font-bold mb-2">
              Phone Number:
            </label>
            <Field
              type="tel"
              id="phoneNumber"
              name="phoneNumber"
              className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
            />
            <ErrorMessage name="phoneNumber" component="p" className="text-red-500 text-xs mt-1" />
          </div>

          {/* Best Time to Call */}
          <div className="mb-4">
            <p className="text-sm font-bold mb-2">Best Time to Call:</p>
            <div className="flex">
              {[
                { label: 'Morning', value: 'morning' },
                { label: 'Noon', value: 'noon' },
                { label: 'Evening', value: 'evening' },
              ].map(({ label, value }) => (
                <label key={value} className="mr-4">
                  <Field
                    type="checkbox"
                    name="bestTimeToCall"
                    value={value}
                    className="mr-2"
                  />
                  {label}
                </label>
              ))}
            </div>
            <ErrorMessage name="bestTimeToCall" component="p" className="text-red-500 text-xs mt-1" />
          </div>

          {/* Can Leave a VoiceMail */}
          <div className="mb-4">
            <p className="text-sm font-bold mb-2">Can leave a voice mail:</p>
            <div className="flex">
              {[
                { label: 'Yes', value: 'yes' },
                { label: 'No', value: 'no' },
              ].map(({ label, value }) => (
                <label key={value} className="mr-4">
                  <Field
                    type="checkbox"
                    name="canLeaveVoiceMail"
                    value={value}
                    className="mr-2"
                  />
                  {label}
                </label>
              ))}
            </div>
            <ErrorMessage name="bestTimeToCall" component="p" className="text-red-500 text-xs mt-1" />
          </div>


          {/* Address Section */}
          <div className="mb-4">
            <label htmlFor="address" className="block text-sm font-bold mb-2">
              Address:
            </label>
            <Field
              type="text"
              id="address"
              name="address"
              className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
              
            />
            <ErrorMessage name="address" component="p" className="text-red-500 text-xs mt-1" />
          </div>

          <div className="mb-4">
              <label htmlFor="city" className="block text-sm font-bold mb-2">
                City:
              </label>
              <Field
                type="text"
                id="city"
                name="city"
                className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
                
              />
              <ErrorMessage name="city" component="p" className="text-red-500 text-xs mt-1" />
            </div>

          <div className="mb-4 flex">
            <div className="w-1/2 mr-2">
              <label htmlFor="postalCode" className="block text-sm font-bold mb-2">
                Postal Code:
              </label>
              <Field
                type="text"
                id="postalCode"
                name="postalCode"
                className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
                
              />
              <ErrorMessage name="postalCode" component="p" className="text-red-500 text-xs mt-1" />
            </div>

            <div className="w-1/2 ml-2">
              <label htmlFor="province" className="block text-sm font-bold mb-2">
                Province:
              </label>
              <Field
                type="text"
                id="province"
                name="province"
                className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
                
              />
              <ErrorMessage name="province" component="p" className="text-red-500 text-xs mt-1" />
            </div>
          </div>

          {/* Email Field */}
          <div className="mb-4">
            <label htmlFor="email" className="block text-sm font-bold mb-2">
              Email:
            </label>
            <Field
              type="email"
              id="email"
              name="email"
              className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
            />
            <ErrorMessage name="email" component="p" className="text-red-500 text-xs mt-1" />
          </div>

          {/* Physician Field*/}
          <div className="mb-4">
            <label htmlFor="email" className="block text-sm font-bold mb-2">
            Physician Name:
            </label>
            <Field
              type="text"
              id="physician"
              name="physician"
              className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
            />
            <ErrorMessage name="email" component="p" className="text-red-500 text-xs mt-1" />
          </div>

          {/* Therapies */}
          <div className="mb-4">
            <p className="text-sm font-bold mb-2">Are you on Multiple Sclerosis/Neuromyelitis Optica Therapies: </p>
            <div className="flex">
              {[
                { label: 'Yes', value: 'yes' },
                { label: 'No', value: 'no' },
              ].map(({ label, value }) => (
                <label key={value} className="mr-4">
                  <Field
                    type="checkbox"
                    name="therapiesCheckbox"
                    value={value}
                    className="mr-2"
                  />
                  {label}
                </label>
              ))}
            </div>
            <ErrorMessage name="therapies" component="p" className="text-red-500 text-xs mt-1" />
          </div>

          {/* If Yes Select */}
          <div className="mb-4">
            <p className="text-sm font-bold mb-2">If Yes, please select from below: </p>
            <div className="flex flex-col">
              {[
                { label: 'Aubagio (Teriflunimide)', value: 'Aubagio (Teriflunimide)' },
                { label: 'Enspryng (Satralizumab)', value: 'Enspryng (Satralizumab)' },
                { label: 'Gilenya (Fingolimod)', value: 'Gilenya (Fingolimod)' },
                { label: 'Kesimpta (Ofatumumab)', value: 'Kesimpta (Ofatumumab)' },
                { label: 'Lemtrada (Alemtuzumab)', value: 'Lemtrada (Alemtuzumab)' },
                { label: 'Mavenclad (Cladribine)', value: 'Mavenclad (Cladribine)' },
                { label: 'Ocrevus (Ocrelizumab)', value: 'Ocrevus (Ocrelizumab)' },
                { label: 'Riximyo/Ruxience (Rituximab)', value: 'Riximyo/Ruxience (Rituximab)' },
                { label: 'Tecfidera (Dimethyl Fumarate)', value: 'Tecfidera (Dimethyl Fumarate)' },
                { label: 'Tysabri (Natalizumab)', value: 'Tysabri (Natalizumab)' },
              ].map(({ label, value }) => (
                <label key={value} className="mr-4 mb-2">
                  <Field
                    type="radio"
                    name="typeOfTherapiesCheckbox"
                    value={value}
                    className="mr-2"
                  />
                  {label}
                </label>
              ))}
              {/* Add "Other" checkbox and text input */}
              <label className="mr-4 mb-2">
                <Field
                  type="radio"
                  name="typeOfTherapiesCheckbox"
                  value="Other"
                  className="mr-2"
                />
                Other
                {/* Conditional rendering of the text input based on the "Other" checkbox */}
                <Field
                  type="text"
                  name="otherTherapy"
                  className="ml-2 pl-2 border rounded-md"
                  placeholder=""
                />
              </label>
            </div>
            <ErrorMessage name="typeOfTherapies" component="p" className="text-red-500 text-xs mt-1" />
          </div>

          {/* Primary Insurance Provider */}
            <div className="mb-4">
              <label htmlFor="primaryInsuranceProvider" className="block text-sm font-bold mb-2">
                Primary Insurance Provider:
              </label>
              <Field
                type="text"
                id="primaryInsuranceProvider"
                name="primaryInsuranceProvider"
                className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
              />
              <ErrorMessage name="primaryInsuranceProvider" component="p" className="text-red-500 text-xs mt-1" />
            </div>

            {/* Insurer Contract # */}
            <div className="mb-4">
              <label htmlFor="insurerContract" className="block text-sm font-bold mb-2">
                Insurer Contract #:
              </label>
              <Field
                type="text"
                id="insurerContract"
                name="insurerContract"
                className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
              />
              <ErrorMessage name="insurerContract" component="p" className="text-red-500 text-xs mt-1" />
            </div>

            {/* Insurer Certificate # */}
            <div className="mb-4">
              <label htmlFor="insurerCertificate" className="block text-sm font-bold mb-2">
                Insurer Certificate #:
              </label>
              <Field
                type="text"
                id="insurerCertificate"
                name="insurerCertificate"
                className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
              />
              <ErrorMessage name="insurerCertificate" component="p" className="text-red-500 text-xs mt-1" />
            </div>





            {/* Patient Consent Text */}
              <div className="mb-4">

              <h3 className="mb-2 font-bold">PATIENT CONSENT TO ENROL IN AND RECEIVE SERVICES FROM SENTREX</h3>
                <p className="text-sm mb-2 whitespace-pre-line">
                {textTermsAndConditions}
                </p>
              </div>




        {/* Signature field */}
      <div className="mb-4">
        <label htmlFor="signature" className="block text-sm font-bold mb-2">
          Patient&apos;s Signature:
        </label>
        <div className="border border-gray-300 rounded-md p-4">
          <SignatureCanvas
            ref={signaturePad}
            penColor="black"
            canvasProps={{
              className: 'w-full h-20',
            }}
          />
        </div>
        <ErrorMessage name="signature" component="p" className="text-red-500 text-xs mt-1" />
      </div>

      {/* Buttons */}
      <div className="flex space-x-4 mb-4">
        <button
          type="button"
          onClick={() => {
            const signatureData = signaturePad.current.toDataURL();
            console.log('Signature Data:', signatureData);
          }}
          className="w-full bg-blue-500 text-white py-2 rounded-md"
        >
          Save Signature
        </button>

        
      </div>


          {/* Date input for Patient Signature */}
          <div className="mb-8 mt-3 flex items-end">
            <label htmlFor="patientSignatureDate" className="block text-sm font-bold mr-2">
              Date of Patient&apos;s Signature:
            </label>
            <div className="relative">
              <Field
                type="date"
                id="patientSignatureDate"
                name="patientSignatureDate"
                className="mt-2 px-3 py-1 border-none focus:outline-none"
              />
              <div className="absolute left-0 right-0 bottom-0 h-px bg-gray-500" />
            </div>

            <ErrorMessage name="patientSignatureDate" component="p" className="text-red-500 text-xs mt-1" />
          </div>







        {/* Printed Name of Patient or Patient’s SDM */}
          <div className="mb-4">
            <label htmlFor="patientPrintedNameOrSDM" className="block text-sm font-bold mb-2">
              Printed Name of Patient or Patient’s SDM:
            </label>
              <Field
                type="text"
                id="patientPrintedNameOrSDM"
                name="patientPrintedNameOrSDM"
                className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
              />
            <ErrorMessage name="patientPrintedNameOrSDM" component="p" className="text-red-500 text-xs mt-1" />
          </div>


          {/* Relationship of SDM to Patient */}
          <div className="mb-4">
            <label htmlFor="relationshipToPatient" className="block text-sm font-bold mb-2">
              Relationship of SDM to Patient:
            </label>
            
              <Field
                type="text"
                id="relationshipToPatient"
                name="relationshipToPatient"
                className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
              />
            <ErrorMessage name="relationshipToPatient" component="p" className="text-red-500 text-xs mt-1" />
          </div>

          {/* Verbal Consent */}
          <div className="mb-4 mt-5">
            
              
              <label className="mr-4 mb-2 text-sm font-bold">
                <Field
                  type="checkbox"
                  name="VerbalConsent"
                  value="VerbalConsent"
                  className="mr-2 "
                />
                Verbal Consent Obtained: I have read the above consent to Patient
                {/* Conditional rendering of the text input based on the "Other" checkbox */}
                
              </label>
            
            <ErrorMessage name="verbalConsentPatientTo" component="p" className="text-red-500 text-xs mt-1" />
          </div>

          
          <div className="mb-4 mt-3">
  {/* Verbal Consent Obtained By */}
  <div className="mb-4">
    <label htmlFor="verbalConsentObtainedBy" className="block text-sm font-bold mr-2">
      Verbal Consent Obtained By:
    </label>
      <Field
        type="text"
        id="verbalConsentObtainedBy"
        name="verbalConsentObtainedBy"
        className="w-full px-3 py-2 border rounded-md focus:outline-none focus:border-blue-500"
      />
    <ErrorMessage name="verbalConsentObtainedBy" component="p" className="text-red-500 text-xs mt-1" />
  </div>
</div>

<div className="mb-4">
  {/* Verbal Consent Date input */}
  <div className="flex items-end">
    <label htmlFor="verbalConsentDate" className="block text-sm font-bold mr-2">
      Date of Verbal Consent:
    </label>
    <div className="relative">
      <Field
        type="date"
        id="verbalConsentDate"
        name="verbalConsentDate"
        className=" px-3 py-1 border-none focus:outline-none"
      />
      <div className="absolute left-0 right-0 bottom-0 h-px bg-gray-500" />
    </div>
    <ErrorMessage name="verbalConsentDate" component="p" className="text-red-500 text-xs mt-1" />
  </div>
</div>





        <button type="submit" className="w-full bg-blue-500 text-white py-2 rounded-md mt-4">
          Submit Consent Form
        </button>

        
      </Form>
    </Formik>
  </div>
);
}

export default App;