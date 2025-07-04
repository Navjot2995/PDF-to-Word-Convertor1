import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import numpy as np
import cv2
import io

class OCRProcessor:
    def __init__(self):
        # Configure Tesseract for better accuracy
        self.config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!@#$%^&*()_+-=[]{}|;:,.<>?/~` '
        
    def preprocess_image(self, image):
        """Preprocess image for better OCR results"""
        try:
            # Convert PIL Image to OpenCV format
            opencv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            
            # Convert to grayscale
            gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
            
            # Apply Gaussian blur to reduce noise
            blurred = cv2.GaussianBlur(gray, (5, 5), 0)
            
            # Apply adaptive thresholding
            thresh = cv2.adaptiveThreshold(
                blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2
            )
            
            # Apply morphological operations to clean up the image
            kernel = np.ones((2, 2), np.uint8)
            cleaned = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
            
            # Convert back to PIL Image
            processed_image = Image.fromarray(cleaned)
            
            return processed_image
            
        except Exception as e:
            print(f"Error preprocessing image: {e}")
            return image
    
    def enhance_image(self, image):
        """Enhance image quality for OCR"""
        try:
            # Enhance contrast
            enhancer = ImageEnhance.Contrast(image)
            image = enhancer.enhance(1.5)
            
            # Enhance sharpness
            enhancer = ImageEnhance.Sharpness(image)
            image = enhancer.enhance(1.2)
            
            # Apply unsharp mask filter
            image = image.filter(ImageFilter.UnsharpMask(radius=2, percent=150, threshold=3))
            
            return image
            
        except Exception as e:
            print(f"Error enhancing image: {e}")
            return image
    
    def extract_text_from_image(self, image):
        """Extract text from a single image using OCR"""
        try:
            # Preprocess the image
            processed_image = self.preprocess_image(image)
            
            # Enhance the image
            enhanced_image = self.enhance_image(processed_image)
            
            # Perform OCR with different PSM modes for better results
            psm_modes = [6, 3, 4, 7, 8, 11, 12, 13]
            best_text = ""
            best_confidence = 0
            
            for psm in psm_modes:
                try:
                    config = f'--oem 3 --psm {psm}'
                    
                    # Extract text
                    text = pytesseract.image_to_string(enhanced_image, config=config)
                    
                    # Get confidence scores
                    data = pytesseract.image_to_data(enhanced_image, config=config, output_type=pytesseract.Output.DICT)
                    confidences = [int(conf) for conf in data['conf'] if int(conf) > 0]
                    
                    if confidences:
                        avg_confidence = sum(confidences) / len(confidences)
                        
                        # Use the result with highest confidence
                        if avg_confidence > best_confidence and len(text.strip()) > 0:
                            best_confidence = avg_confidence
                            best_text = text
                
                except Exception as e:
                    continue
            
            # If no good result found, try with original image
            if not best_text.strip():
                best_text = pytesseract.image_to_string(image, config='--oem 3 --psm 6')
            
            return best_text.strip()
            
        except Exception as e:
            print(f"Error extracting text from image: {e}")
            return ""
    
    def process_images(self, image_data):
        """Process multiple images and combine extracted text"""
        try:
            combined_text = ""
            
            for img_info in image_data:
                page_num = img_info['page_number']
                image = img_info['image']
                
                # Add page separator
                if page_num > 1:
                    combined_text += f"\n\n--- Page {page_num} ---\n\n"
                
                # Extract text from current page
                page_text = self.extract_text_from_image(image)
                
                if page_text:
                    combined_text += page_text
                else:
                    combined_text += f"[No text detected on page {page_num}]"
                
                combined_text += "\n"
            
            return combined_text.strip()
            
        except Exception as e:
            raise Exception(f"Error processing images: {e}")
    
    def detect_handwriting(self, image):
        """Detect if image contains handwritten text (basic heuristic)"""
        try:
            # Convert to grayscale
            gray = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2GRAY)
            
            # Apply edge detection
            edges = cv2.Canny(gray, 50, 150)
            
            # Find contours
            contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            
            # Analyze contour characteristics
            irregular_contours = 0
            for contour in contours:
                # Calculate contour area and perimeter
                area = cv2.contourArea(contour)
                if area > 100:  # Filter small noise
                    perimeter = cv2.arcLength(contour, True)
                    if perimeter > 0:
                        circularity = 4 * np.pi * area / (perimeter * perimeter)
                        # Handwritten text tends to have more irregular shapes
                        if circularity < 0.3:
                            irregular_contours += 1
            
            # If more than 30% of contours are irregular, likely handwritten
            if len(contours) > 0 and irregular_contours / len(contours) > 0.3:
                return True
            
            return False
            
        except Exception as e:
            print(f"Error detecting handwriting: {e}")
            return False
