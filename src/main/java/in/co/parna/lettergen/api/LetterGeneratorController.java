package in.co.parna.lettergen.api;

import com.lowagie.text.DocumentException;
import in.co.parna.lettergen.dto.LetterGeneratorRequestDto;
import in.co.parna.lettergen.service.LetterGeneratorService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@Api(value = "Letter Generator", tags = "Letter Generator")
@RestController("api/v1/letter-gen")
public class LetterGeneratorController {

    @Autowired
    private LetterGeneratorService letterGeneratorService;

    @ApiOperation(value = "Letter Generator with provided information", tags = "Letter Generator")
    @RequestMapping(method = RequestMethod.PUT)
    public void generateLetters(@RequestBody LetterGeneratorRequestDto letterGeneratorRequestDto) throws IOException, DocumentException {

        letterGeneratorService.generateLetters(letterGeneratorRequestDto);
    }
}
